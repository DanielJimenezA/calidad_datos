"""Visor local de auditoría Parquet EDS. Escucha exclusivamente en 127.0.0.1."""
import argparse
import html
import json
import csv
import io
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse, urlencode
import duckdb

CAMPOS = ['MODULO','ENTIDAD','CLUES','ARCHIVO_ORIGEN','FILA_ORIGEN','CLAVE_REGISTRO','REGLA_ID','VARIABLE','VALOR_OBSERVADO','MOTIVO']
FILTROS = ['MODULO','ENTIDAD','CLUES','REGLA_ID']

def obtener(ruta, args, exportar=False):
    conn = duckdb.connect(':memory:')
    patron = str(ruta.resolve() / '*_Resultados.parquet')
    archivos = list(ruta.glob('*_Resultados.parquet'))
    if not archivos:
        raise FileNotFoundError('No se encontraron archivos *_Resultados.parquet en la carpeta privada.')
    escaped_patron = patron.replace("'", "''")
    conn.execute(f"CREATE VIEW audit AS SELECT * FROM read_parquet('{escaped_patron}', union_by_name=true)")
    columnas = {x[0] for x in conn.execute('DESCRIBE audit').fetchall()}
    condiciones = []
    parametros = []
    for campo in FILTROS:
        valor = args.get(campo, [''])[0].strip()
        if valor and campo in columnas:
            condiciones.append(f'CAST("{campo}" AS VARCHAR) ILIKE ?')
            parametros.append('%' + valor + '%')
    busqueda = args.get('q', [''])[0].strip()
    if busqueda:
        campos = [x for x in CAMPOS if x in columnas]
        condiciones.append('(' + ' OR '.join(f'CAST("{c}" AS VARCHAR) ILIKE ?' for c in campos) + ')')
        parametros.extend(['%' + busqueda + '%'] * len(campos))
    where = ' WHERE ' + ' AND '.join(condiciones) if condiciones else ''
    total = conn.execute('SELECT count(*) FROM audit' + where, parametros).fetchone()[0]
    try: pagina = max(1, int(args.get('pagina', ['1'])[0]))
    except ValueError: pagina = 1
    columnas_vista = [c for c in CAMPOS if c in columnas]
    cols_sql = ','.join('"' + c + '"' for c in columnas_vista)
    if exportar:
        # Límite de seguridad para exportación interactiva; no limita el Parquet.
        sql = f'SELECT {cols_sql} FROM audit' + where + ' LIMIT 100000'
        rows = conn.execute(sql, parametros).fetchall()
    else:
        sql = f'SELECT {cols_sql} FROM audit' + where + ' LIMIT 100 OFFSET ?'
        rows = conn.execute(sql, parametros + [(pagina - 1) * 100]).fetchall()
    conn.close()
    return total, pagina, columnas_vista, rows

def html_pagina(ruta, args):
    total, pagina, columnas, rows = obtener(ruta, args)
    esc = lambda x: html.escape('' if x is None else str(x), quote=True)
    inputs = ''.join(f'<label>{c}<input name="{c}" value="{esc(args.get(c,[""])[0])}"></label>' for c in FILTROS)
    inputs += f'<label>Buscar<input name="q" value="{esc(args.get("q",[""])[0])}"></label>'
    th = ''.join(f'<th>{esc(c)}</th>' for c in columnas)
    trs = ''.join('<tr>' + ''.join(f'<td>{esc(v)}</td>' for v in row) + '</tr>' for row in rows)
    base = {k:v[0] for k,v in args.items() if k != 'pagina' and v}
    prev = '/?' + urlencode({**base, 'pagina': max(1,pagina-1)})
    nxt = '/?' + urlencode({**base, 'pagina': pagina+1})
    csv_url = '/exportar.csv?' + urlencode(base)
    return f"""<!doctype html><html lang="es"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Auditoría EDS privada</title>
<style>body{{font:15px system-ui;margin:24px;color:#10312B}}header{{background:#611232;color:white;padding:18px;border-radius:8px}}form{{display:flex;flex-wrap:wrap;gap:12px;margin:18px 0}}label{{display:grid;gap:5px}}input{{padding:9px;border:1px solid #aaa;border-radius:5px}}button{{background:#235B4E;color:white;border:0;padding:10px 20px;border-radius:5px}}.table{{overflow:auto}}table{{border-collapse:collapse;width:100%}}th,td{{border-bottom:1px solid #ddd;padding:8px;text-align:left;white-space:nowrap}}th{{background:#e8efe9}}a{{margin:12px}}</style>
<header><h2>EDS · Auditoría privada de incumplimientos</h2><small>Visor local · datos consultados directamente desde Parquet</small></header>
<form method="get">{inputs}<button>Filtrar</button></form><p><a href="{esc(csv_url)}">Exportar CSV filtrado (máximo 100,000 incidencias)</a></p><p><strong>{total:,}</strong> incumplimientos encontrados · Página {pagina}</p>
<div class="table"><table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table></div>
<nav><a href="{esc(prev)}">Anterior</a><a href="{esc(nxt)}">Siguiente</a></nav></html>"""

class Handler(BaseHTTPRequestHandler):
    carpeta = None
    def do_GET(self):
        try:
            parsed = urlparse(self.path)
            if parsed.path == '/exportar.csv':
                _, _, cols, rows = obtener(self.carpeta, parse_qs(parsed.query), exportar=True)
                buffer = io.StringIO()
                writer = csv.writer(buffer)
                writer.writerow(cols)
                writer.writerows(rows)
                data = ('\ufeff' + buffer.getvalue()).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/csv; charset=utf-8')
                self.send_header('Content-Disposition', 'attachment; filename=incidencias_filtradas.csv')
            elif parsed.path == '/':
                html_out = html_pagina(self.carpeta, parse_qs(parsed.query))
                data = html_out.encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
            else:
                self.send_error(404)
                return
        except Exception as exc:
            data = ('Error: ' + html.escape(str(exc))).encode('utf-8')
            self.send_response(500)
        self.send_header('Cache-Control','no-store')
        self.send_header('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'")
        self.end_headers();self.wfile.write(data)

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--carpeta',required=True,help='Ruta a auditoria_privada')
    parser.add_argument('--puerto',type=int,default=8765)
    args=parser.parse_args()
    Handler.carpeta=Path(args.carpeta).resolve()
    print(f'Visor privado: http://127.0.0.1:{args.puerto}')
    ThreadingHTTPServer(('127.0.0.1',args.puerto),Handler).serve_forever()
