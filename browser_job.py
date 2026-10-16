"""Bounded worker communication; Playwright runs in its own process on Windows."""
import json
import queue
import subprocess
import sys
import threading
from pathlib import Path


def browser_events(numbers, worker=None, idle_timeout=75, mode="local"):
    worker = Path(worker) if worker else Path(__file__).with_name('dpd_browser.py')
    process = subprocess.Popen([sys.executable,'-u',str(worker)],stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,text=True,encoding='utf-8',bufsize=1)
    lines = queue.Queue()

    def read_lines():
        try:
            for line in process.stdout:
                lines.put(line)
        finally:
            lines.put(None)

    reader = threading.Thread(target=read_lines,daemon=True)
    reader.start()
    completed = False
    reported_error = False
    try:
        process.stdin.write(json.dumps({'numbers':numbers,'mode':mode}))
        process.stdin.close()
        while True:
            try:
                line = lines.get(timeout=idle_timeout)
            except queue.Empty:
                yield {'type':'error','message':'Браузер не отвечает более 75 секунд. Уже полученные строки сохранены; оставшиеся не проверены.'}
                reported_error = True
                break
            if line is None:
                break
            try:
                event = json.loads(line)
                if not isinstance(event,dict) or event.get('type') not in ('row','done','error'):
                    raise ValueError('Unknown event')
            except (ValueError,TypeError):
                yield {'type':'error','message':'Ошибка обмена с браузером. Полученные строки сохранены.'}
                reported_error = True
                break
            completed |= event['type'] == 'done'
            reported_error |= event['type'] == 'error'
            yield event
        if not completed and not reported_error:
            yield {'type':'error','message':'Проверка прервана до завершения. Полученные строки сохранены; оставшиеся не проверены.'}
    finally:
        if not process.stdin.closed:
            process.stdin.close()
        if process.poll() is None:
            try:
                process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                process.terminate()
                try:
                    process.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=3)
        reader.join(timeout=2)
        if not reader.is_alive():
            process.stdout.close()
