"""本地 PDF 转 Word：启动后在浏览器打开 http://127.0.0.1:端口/

只监听本机 127.0.0.1，用 Python 自带的 http.server，文件不会离开这台电脑。
也可以不开界面直接批量转换：python app.py --batch 文件夹
"""
import io
import json
import os
import pathlib
import socket
import subprocess
import sys
import threading
import traceback
import urllib.parse
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import converter

ROOT = pathlib.Path(__file__).resolve().parent
OUT = ROOT / '转换结果'
UI = (ROOT / 'ui.html').read_bytes()
LOCK = threading.Lock()   # 一次只转一份，避免同时占满内存


def safe_name(name):
    name = pathlib.Path(name).name
    return ''.join(c for c in name if c not in '\\/:*?"<>|\r\n').strip() or '未命名.pdf'


def convert_bytes(data, filename):
    OUT.mkdir(exist_ok=True)
    tmp = OUT / f'.{os.getpid()}-{threading.get_ident()}.pdf'
    stem = pathlib.Path(safe_name(filename)).stem
    target = OUT / f'{stem}.docx'
    n = 2
    while target.exists():
        target = OUT / f'{stem} ({n}).docx'
        n += 1
    try:
        tmp.write_bytes(data)
        with LOCK:
            pages = converter.convert(str(tmp), str(target))
        return {'ok': True, 'status': '成功', 'pages': pages, 'output': target.name}
    except converter.ScannedPDF as e:
        return {'ok': False, 'status': '扫描件已跳过', 'error': str(e)}
    except Exception as e:  # 单个文件失败不影响其他文件
        target.unlink(missing_ok=True)
        traceback.print_exc()
        text = f'{type(e).__name__}: {e}'
        if 'password' in text.lower() or 'encrypt' in text.lower():
            msg = 'PDF 设置了密码，请先解除密码保护'
        elif any(k in text for k in ('EOF', 'PSSyntax', 'PDFSyntax', 'PdfminerException', 'PdfiumError')):
            msg = f'文件损坏或不完整，无法读取（{text}）'
        else:
            msg = f'转换出错（{text}）'
        return {'ok': False, 'status': '失败', 'error': msg}
    finally:
        tmp.unlink(missing_ok=True)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _local_only(self):
        # 只接受本页面发来的请求，防止其他网站借浏览器访问本工具
        host = self.headers.get('Host', '')
        origin = self.headers.get('Origin')
        ok = host.startswith(('127.0.0.1:', 'localhost:')) and (origin is None or origin == f'http://{host}')
        if not ok:
            self._send(403, {'error': 'forbidden'})
        return ok

    def do_GET(self):
        if not self._local_only():
            return
        url = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(url.query)
        if url.path == '/':
            self._send(200, UI, 'text/html; charset=utf-8', {
                'Content-Security-Policy': "default-src 'self'; style-src 'unsafe-inline'; "
                                           "script-src 'unsafe-inline'; img-src data:; connect-src 'self'"})
        elif url.path == '/download':
            names = q.get('name', [])
            files = [OUT / safe_name(n) for n in names if (OUT / safe_name(n)).is_file()]
            if not files:
                return self._send(404, {'error': '文件不存在'})
            if len(files) == 1 and 'zip' not in q:
                f = files[0]
                return self._send(200, f.read_bytes(),
                                  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                  {'Content-Disposition': 'attachment; filename="document.docx"; filename*=UTF-8\'\'' + urllib.parse.quote(f.name)})
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
                for f in files:
                    z.write(f, f.name)
            self._send(200, buf.getvalue(), 'application/zip',
                       {'Content-Disposition': 'attachment; filename="result.zip"; filename*=UTF-8\'\'' + urllib.parse.quote('转换结果.zip')})
        else:
            self._send(404, {'error': 'not found'})

    def do_POST(self):
        if not self._local_only():
            return
        url = urllib.parse.urlparse(self.path)
        if url.path == '/convert':
            name = urllib.parse.parse_qs(url.query).get('name', ['未命名.pdf'])[0]
            length = int(self.headers.get('Content-Length', 0))
            data = self.rfile.read(length)
            if not data.startswith(b'%PDF'):
                return self._send(200, {'ok': False, 'status': '失败', 'error': '这不是 PDF 文件'})
            self._send(200, convert_bytes(data, name))
        elif url.path == '/open':
            OUT.mkdir(exist_ok=True)
            if sys.platform == 'darwin':
                subprocess.Popen(['open', str(OUT)])
            elif sys.platform.startswith('win'):
                os.startfile(str(OUT))  # noqa
            else:
                subprocess.Popen(['xdg-open', str(OUT)])
            self._send(200, {'ok': True})
        else:
            self._send(404, {'error': 'not found'})


def free_port(start=8765):
    for port in range(start, start + 50):
        with socket.socket() as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    raise RuntimeError('找不到可用端口')


def batch(folder):
    folder = pathlib.Path(folder)
    pdfs = sorted(p for p in folder.iterdir() if p.suffix.lower() == '.pdf')
    lines = []
    for p in pdfs:
        r = convert_bytes(p.read_bytes(), p.name)
        lines.append(f"{r['status']}：{p.name}" + (f" → {r['output']}（{r['pages']} 页）" if r['ok'] else f"（{r['error']}）"))
        print(lines[-1])
    (OUT / '转换报告.txt').write_text('\n'.join(lines), encoding='utf-8')


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == '--batch':
        return batch(sys.argv[2])
    port = free_port()
    server = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    url = f'http://127.0.0.1:{port}/'
    print(f'工具已启动：{url}\n关闭这个窗口即可退出。')
    if '--no-browser' not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main()
