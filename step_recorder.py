import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import mss
from PIL import Image, ImageDraw, ImageTk
from pynput import mouse, keyboard
import base64
from io import BytesIO
import time
import os
import sys
import webbrowser

try:
    import uiautomation as auto
    HAS_UI_AUTO = True
except ImportError:
    HAS_UI_AUTO = False

CTRL_TYPE_ZH = {
    'Button': '按鈕',
    'CheckBox': '勾選框',
    'ComboBox': '下拉選單',
    'Edit': '輸入框',
    'Hyperlink': '連結',
    'Image': '圖片',
    'ListItem': '清單項目',
    'List': '清單',
    'Menu': '選單',
    'MenuBar': '選單列',
    'MenuItem': '選單項目',
    'ProgressBar': '進度條',
    'RadioButton': '選項按鈕',
    'ScrollBar': '捲軸',
    'Slider': '滑桿',
    'StatusBar': '狀態列',
    'Tab': '標籤頁',
    'TabItem': '標籤',
    'Text': '文字',
    'TitleBar': '標題列',
    'ToolBar': '工具列',
    'Tree': '樹狀結構',
    'TreeItem': '樹狀項目',
    'Window': '視窗',
    'Pane': '面板',
    'Group': '群組',
    'Header': '標題',
    'Table': '表格',
    'Document': '文件',
    'Custom': '元件',
}


MODIFIER_KEYS = {
    keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r,
    keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r,
    keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r,
    keyboard.Key.cmd, keyboard.Key.cmd_l, keyboard.Key.cmd_r,
}

STANDALONE_KEYS = {
    keyboard.Key.enter: 'Enter',
    keyboard.Key.delete: 'Delete',
    keyboard.Key.backspace: 'Backspace',
    keyboard.Key.tab: 'Tab',
    keyboard.Key.esc: 'Esc',
    keyboard.Key.space: '空白鍵',
    keyboard.Key.f1: 'F1', keyboard.Key.f2: 'F2', keyboard.Key.f3: 'F3',
    keyboard.Key.f4: 'F4', keyboard.Key.f5: 'F5', keyboard.Key.f6: 'F6',
    keyboard.Key.f7: 'F7', keyboard.Key.f8: 'F8', keyboard.Key.f9: 'F9',
    keyboard.Key.f10: 'F10', keyboard.Key.f11: 'F11', keyboard.Key.f12: 'F12',
    keyboard.Key.home: 'Home', keyboard.Key.end: 'End',
    keyboard.Key.page_up: 'Page Up', keyboard.Key.page_down: 'Page Down',
}


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)


def find_chinese_font():
    candidates = [
        r'C:\Windows\Fonts\msjh.ttc',
        r'C:\Windows\Fonts\mingliu.ttc',
        r'C:\Windows\Fonts\kaiu.ttf',
        r'C:\Windows\Fonts\simsun.ttc',
        r'C:\Windows\Fonts\msgothic.ttc',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


class StepData:
    def __init__(self, step_id, description, image_b64, timestamp):
        self.step_id = step_id
        self.description = description
        self.image_b64 = image_b64
        self.timestamp = timestamp


class Recorder:
    def __init__(self):
        self.steps = []
        self.selected_monitor_idx = 0
        self.monitors = []
        self.is_recording = False
        self.is_paused = False
        self.listener = None
        self.kb_listener = None
        self._lock = threading.Lock()
        self._modifiers = set()
        self._last_shortcut = ('', 0)  # (combo, timestamp) 防止重複觸發
        self.on_step_added = None

    def load_monitors(self):
        with mss.mss() as sct:
            self.monitors = list(sct.monitors[1:])

    def get_selected_monitor(self):
        return self.monitors[self.selected_monitor_idx]

    def is_in_selected_monitor(self, x, y):
        m = self.get_selected_monitor()
        return (m['left'] <= x < m['left'] + m['width'] and
                m['top'] <= y < m['top'] + m['height'])

    def get_active_window_name(self):
        if not HAS_UI_AUTO:
            return ''
        try:
            ctrl = auto.GetFocusedControl()
            if ctrl:
                window = ctrl.GetTopLevelControl()
                return window.Name if window else ''
        except Exception:
            pass
        return ''

    def format_shortcut(self, key):
        parts = []
        if any(k in self._modifiers for k in (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r)):
            parts.append('Ctrl')
        if any(k in self._modifiers for k in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r)):
            parts.append('Shift')
        if any(k in self._modifiers for k in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r)):
            parts.append('Alt')
        if any(k in self._modifiers for k in (keyboard.Key.cmd, keyboard.Key.cmd_l, keyboard.Key.cmd_r)):
            parts.append('Win')
        if key in STANDALONE_KEYS:
            parts.append(STANDALONE_KEYS[key])
        elif hasattr(key, 'char') and key.char:
            parts.append(key.char.upper())
        else:
            parts.append(str(key).replace('Key.', '').upper())
        return '+'.join(parts)

    def capture_shortcut(self, combo):
        try:
            window_name = self.get_active_window_name()
            parts = []
            if window_name:
                parts.append(f"【{window_name}】")
            parts.append(f"按下快捷鍵 {combo}")
            description = '　'.join(parts)

            m = self.get_selected_monitor()
            with mss.mss() as sct:
                screenshot = sct.grab(m)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            buf = BytesIO()
            w, h = img.size
            if w > 1600:
                ratio = 1600 / w
                img = img.resize((int(w*ratio), int(h*ratio)), Image.LANCZOS)
            img.save(buf, format='JPEG', quality=85)
            img_b64 = base64.b64encode(buf.getvalue()).decode()

            step = StepData(
                step_id=len(self.steps) + 1,
                description=description,
                image_b64=img_b64,
                timestamp=time.strftime('%H:%M:%S')
            )
            with self._lock:
                self.steps.append(step)
            if self.on_step_added:
                self.on_step_added(len(self.steps))
        except Exception as e:
            print(f"Shortcut capture error: {e}")

    def on_key_press(self, key):
        if not self.is_recording or self.is_paused:
            return
        if key in MODIFIER_KEYS:
            self._modifiers.add(key)
            return
        has_modifier = bool(self._modifiers)
        is_standalone = key in STANDALONE_KEYS
        if not has_modifier and not is_standalone:
            return
        combo = self.format_shortcut(key)
        now = time.time()
        if combo == self._last_shortcut[0] and now - self._last_shortcut[1] < 0.5:
            return
        self._last_shortcut = (combo, now)
        threading.Thread(target=self.capture_shortcut, args=(combo,), daemon=True).start()

    def on_key_release(self, key):
        self._modifiers.discard(key)

    def get_element_description(self, x, y):
        if not HAS_UI_AUTO:
            return f"點擊位置 ({x}, {y})"
        try:
            ctrl = auto.ControlFromPoint(x, y)
            if ctrl:
                name = ctrl.Name or ''
                ctrl_type_en = (ctrl.ControlTypeName or '').replace('Control', '').strip()
                ctrl_type = CTRL_TYPE_ZH.get(ctrl_type_en, ctrl_type_en)
                window = ctrl.GetTopLevelControl()
                window_name = window.Name if window else ''
                parts = []
                if window_name:
                    parts.append(f"【{window_name}】")
                if name:
                    parts.append(f"點擊{ctrl_type}「{name}」")
                elif ctrl_type:
                    parts.append(f"點擊{ctrl_type}")
                return '　'.join(parts) if parts else f"點擊位置 ({x}, {y})"
        except Exception:
            pass
        return f"點擊位置 ({x}, {y})"

    def capture_step(self, x, y):
        try:
            description = self.get_element_description(x, y)
            m = self.get_selected_monitor()
            with mss.mss() as sct:
                screenshot = sct.grab(m)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')

            rel_x = x - m['left']
            rel_y = y - m['top']
            draw = ImageDraw.Draw(img)
            r = 22
            draw.ellipse([rel_x-r, rel_y-r, rel_x+r, rel_y+r], outline='red', width=3)
            draw.ellipse([rel_x-4, rel_y-4, rel_x+4, rel_y+4], fill='red')

            buf = BytesIO()
            w, h = img.size
            if w > 1600:
                ratio = 1600 / w
                img = img.resize((int(w*ratio), int(h*ratio)), Image.LANCZOS)
            img.save(buf, format='JPEG', quality=85)
            img_b64 = base64.b64encode(buf.getvalue()).decode()

            step = StepData(
                step_id=len(self.steps) + 1,
                description=description,
                image_b64=img_b64,
                timestamp=time.strftime('%H:%M:%S')
            )
            with self._lock:
                self.steps.append(step)
            if self.on_step_added:
                self.on_step_added(len(self.steps))
        except Exception as e:
            print(f"Capture error: {e}")

    def on_click(self, x, y, button, pressed):
        if not pressed or not self.is_recording or self.is_paused:
            return
        if button != mouse.Button.left:
            return
        if not self.is_in_selected_monitor(x, y):
            return
        threading.Thread(target=self.capture_step, args=(x, y), daemon=True).start()

    def start(self):
        self.is_recording = True
        self.is_paused = False
        self.listener = mouse.Listener(on_click=self.on_click)
        self.listener.start()
        self.kb_listener = keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release
        )
        self.kb_listener.start()

    def pause(self):
        self.is_paused = True

    def resume(self):
        self.is_paused = False

    def stop(self):
        self.is_recording = False
        if self.listener:
            self.listener.stop()
            self.listener = None
        if self.kb_listener:
            self.kb_listener.stop()
            self.kb_listener = None


class SetupWindow:
    def __init__(self, recorder, on_start):
        self.recorder = recorder
        self.on_start = on_start
        self.win = tk.Tk()
        self.win.title("教學步驟記錄器")
        self.win.resizable(False, False)
        self._build()

    def _build(self):
        frame = ttk.Frame(self.win, padding=28)
        frame.pack()
        ttk.Label(frame, text="教學步驟記錄器", font=('Microsoft JhengHei', 16, 'bold')).pack(pady=(0, 6))
        ttk.Label(frame, text="選擇要記錄的螢幕", font=('Microsoft JhengHei', 11)).pack(pady=(0, 14))
        self.var = tk.IntVar(value=0)
        for i, m in enumerate(self.recorder.monitors):
            label = f"螢幕 {i+1}　（{m['width']} × {m['height']}）"
            ttk.Radiobutton(frame, text=label, variable=self.var, value=i).pack(anchor='w', pady=3)
        ttk.Button(frame, text="開始記錄", command=self._start, width=20).pack(pady=(20, 0))

    def _start(self):
        self.recorder.selected_monitor_idx = self.var.get()
        self.win.withdraw()
        self.on_start()

    def run(self):
        self.win.mainloop()


class RecordingOverlay:
    def __init__(self, recorder, on_stop):
        self.recorder = recorder
        self.on_stop = on_stop
        self._paused = False
        self.win = tk.Toplevel()
        self.win.title("記錄中")
        self.win.attributes('-topmost', True)
        self.win.attributes('-alpha', 0.75)
        self.win.resizable(False, False)
        self.win.protocol("WM_DELETE_WINDOW", self._stop)
        self._build()
        self._update()

    def _build(self):
        frame = ttk.Frame(self.win, padding=16)
        frame.pack()
        self.status_label = ttk.Label(frame, text="● 記錄中", foreground='red',
                                      font=('Microsoft JhengHei', 11, 'bold'))
        self.status_label.pack()
        self.count_label = ttk.Label(frame, text="已記錄 0 個步驟",
                                     font=('Microsoft JhengHei', 10))
        self.count_label.pack(pady=6)
        btn_row = ttk.Frame(frame)
        btn_row.pack()
        self.pause_btn = ttk.Button(btn_row, text="暫停", command=self._toggle_pause, width=8)
        self.pause_btn.pack(side='left', padx=(0, 6))
        ttk.Button(btn_row, text="停止記錄", command=self._stop, width=8).pack(side='left')

    def _toggle_pause(self):
        self._paused = not self._paused
        if self._paused:
            self.recorder.pause()
            self.pause_btn.config(text="繼續")
            self.status_label.config(text="⏸ 已暫停", foreground='#888')
        else:
            self.recorder.resume()
            self.pause_btn.config(text="暫停")
            self.status_label.config(text="● 記錄中", foreground='red')

    def _update(self):
        if self.win.winfo_exists():
            self.count_label.config(text=f"已記錄 {len(self.recorder.steps)} 個步驟")
            self.win.after(400, self._update)

    def _stop(self):
        self.recorder.stop()
        self.win.destroy()
        self.on_stop()


class EditorWindow:
    def __init__(self, recorder):
        self.recorder = recorder
        self.win = tk.Toplevel()
        self.win.title("編輯步驟")
        self.win.geometry("900x660")
        self._photos = []
        self._build()
        self._render_steps()

    def _build(self):
        top = ttk.Frame(self.win, padding=(14, 10))
        top.pack(fill='x', side='top')
        self.count_label = ttk.Label(top, text="", font=('Microsoft JhengHei', 11))
        self.count_label.pack(side='left')

        btn_frame = ttk.Frame(top)
        btn_frame.pack(side='right')
        ttk.Button(btn_frame, text="匯出 PDF", command=self._export_pdf).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text="匯出 HTML", command=self._export_html).pack(side='left')

        ttk.Separator(self.win, orient='horizontal').pack(fill='x')
        outer = ttk.Frame(self.win)
        outer.pack(fill='both', expand=True)
        self.canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient='vertical', command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind('<Configure>', lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox('all')))
        self.canvas_win = self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(
            self.canvas_win, width=e.width))
        self.canvas.bind_all('<MouseWheel>', lambda e: self.canvas.yview_scroll(
            int(-1*(e.delta/120)), 'units'))
        scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)

    def _render_steps(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self._photos = []
        self.count_label.config(text=f"共 {len(self.recorder.steps)} 個步驟")
        for i, step in enumerate(self.recorder.steps):
            self._render_card(i, step)

    def _render_card(self, i, step):
        card = ttk.Frame(self.inner, relief='groove', padding=14)
        card.pack(fill='x', padx=12, pady=6)
        header = ttk.Frame(card)
        header.pack(fill='x')
        ttk.Label(header, text=f"步驟 {i+1}", font=('Microsoft JhengHei', 11, 'bold')).pack(side='left')
        ttk.Label(header, text=step.timestamp, foreground='#888').pack(side='left', padx=10)
        ttk.Button(header, text="刪除", command=lambda idx=i: self._delete(idx)).pack(side='right')
        var = tk.StringVar(value=step.description)
        entry = ttk.Entry(card, textvariable=var, font=('Microsoft JhengHei', 10))
        entry.pack(fill='x', pady=(8, 10))
        entry.bind('<FocusOut>', lambda e, idx=i, v=var: self._update_desc(idx, v.get()))
        entry.bind('<Return>', lambda e, idx=i, v=var: self._update_desc(idx, v.get()))
        try:
            img = Image.open(BytesIO(base64.b64decode(step.image_b64)))
            img.thumbnail((820, 460), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._photos.append(photo)
            ttk.Label(card, image=photo).pack(anchor='w')
        except Exception as e:
            ttk.Label(card, text=f"（圖片載入失敗：{e}）", foreground='red').pack()

    def _delete(self, idx):
        del self.recorder.steps[idx]
        self._render_steps()

    def _update_desc(self, idx, text):
        if 0 <= idx < len(self.recorder.steps):
            self.recorder.steps[idx].description = text

    def _export_html(self):
        if not self.recorder.steps:
            messagebox.showwarning("注意", "目前沒有任何步驟可以匯出。")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension='.html',
            filetypes=[('HTML 檔案', '*.html')],
            initialfile='教學步驟記錄',
            title='匯出為'
        )
        if not filepath:
            return
        steps_html = ''
        for i, step in enumerate(self.recorder.steps):
            steps_html += f'''
    <div class="step">
      <div class="step-header">
        <div class="step-num">{i+1}</div>
        <div class="step-desc">{step.description}</div>
        <div class="step-time">{step.timestamp}</div>
      </div>
      <img src="data:image/jpeg;base64,{step.image_b64}" alt="步驟 {i+1}">
    </div>'''
        html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>教學步驟記錄</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; background: #f5f5f5; color: #333; }}
    .container {{ max-width: 980px; margin: 0 auto; padding: 36px 20px; }}
    h1 {{ font-size: 24px; margin-bottom: 6px; }}
    .meta {{ color: #888; font-size: 13px; margin-bottom: 36px; }}
    .step {{ background: white; border-radius: 10px; padding: 22px; margin-bottom: 24px;
             box-shadow: 0 1px 5px rgba(0,0,0,0.09); }}
    .step-header {{ display: flex; align-items: center; gap: 12px; margin-bottom: 14px; }}
    .step-num {{ background: #3B82F6; color: white; border-radius: 50%;
                  width: 34px; height: 34px; display: flex; align-items: center;
                  justify-content: center; font-weight: bold; font-size: 14px; flex-shrink: 0; }}
    .step-desc {{ font-size: 15px; flex: 1; }}
    .step-time {{ color: #bbb; font-size: 12px; flex-shrink: 0; }}
    img {{ max-width: 100%; border-radius: 6px; border: 1px solid #e8e8e8; }}
  </style>
</head>
<body>
  <div class="container">
    <h1>教學步驟記錄</h1>
    <p class="meta">共 {len(self.recorder.steps)} 個步驟　｜　匯出時間：{time.strftime('%Y-%m-%d %H:%M')}</p>
    {steps_html}
  </div>
</body>
</html>'''
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        messagebox.showinfo("完成", f"已匯出：\n{filepath}")
        webbrowser.open(filepath)

    def _export_pdf(self):
        if not self.recorder.steps:
            messagebox.showwarning("注意", "目前沒有任何步驟可以匯出。")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension='.pdf',
            filetypes=[('PDF 檔案', '*.pdf')],
            initialfile='教學步驟記錄',
            title='匯出為'
        )
        if not filepath:
            return
        try:
            from fpdf import FPDF
        except ImportError:
            messagebox.showerror("錯誤", "缺少 PDF 套件，請確認程式已正確安裝。")
            return

        font_path = find_chinese_font()

        class PDF(FPDF):
            pass

        pdf = PDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        if font_path:
            try:
                pdf.add_font('CJK', fname=font_path)
                font_name = 'CJK'
            except Exception:
                font_name = 'Helvetica'
        else:
            font_name = 'Helvetica'

        # Title page
        pdf.add_page()
        pdf.set_font(font_name, size=22)
        pdf.ln(20)
        pdf.cell(0, 12, '教學步驟記錄', new_x='LMARGIN', new_y='NEXT', align='C')
        pdf.set_font(font_name, size=11)
        pdf.set_text_color(140, 140, 140)
        pdf.cell(0, 8,
                 f'共 {len(self.recorder.steps)} 個步驟  |  {time.strftime("%Y-%m-%d %H:%M")}',
                 new_x='LMARGIN', new_y='NEXT', align='C')
        pdf.set_text_color(0, 0, 0)

        temp_dir = os.environ.get('TEMP', os.path.expanduser('~'))

        for i, step in enumerate(self.recorder.steps):
            pdf.add_page()

            # Step number badge area
            pdf.set_fill_color(59, 130, 246)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font(font_name, size=13)
            pdf.cell(0, 10, f'  步驟 {i+1}', new_x='LMARGIN', new_y='NEXT', fill=True)
            pdf.set_text_color(0, 0, 0)

            # Description
            pdf.set_font(font_name, size=11)
            pdf.ln(3)
            pdf.multi_cell(0, 7, step.description)

            pdf.set_font(font_name, size=9)
            pdf.set_text_color(160, 160, 160)
            pdf.cell(0, 6, step.timestamp, new_x='LMARGIN', new_y='NEXT')
            pdf.set_text_color(0, 0, 0)
            pdf.ln(4)

            # Image
            try:
                img_data = base64.b64decode(step.image_b64)
                img = Image.open(BytesIO(img_data))
                tmp_path = os.path.join(temp_dir, f'_step_tmp_{i}.jpg')
                img.save(tmp_path, 'JPEG', quality=85)

                page_w = pdf.w - pdf.l_margin - pdf.r_margin
                img_w, img_h = img.size
                display_w = page_w
                display_h = img_h * (page_w / img_w)

                remaining = pdf.h - pdf.get_y() - pdf.b_margin
                if display_h > remaining and display_h > 20:
                    ratio = remaining / display_h
                    display_w *= ratio
                    display_h = remaining

                pdf.image(tmp_path, x=pdf.l_margin, w=display_w, h=display_h)
                os.remove(tmp_path)
            except Exception:
                pdf.cell(0, 8, '（圖片載入失敗）', new_x='LMARGIN', new_y='NEXT')

        pdf.output(filepath)
        messagebox.showinfo("完成", f"已匯出：\n{filepath}")
        os.startfile(filepath)


class App:
    def __init__(self):
        self.recorder = Recorder()
        self.recorder.load_monitors()
        self.setup_win = SetupWindow(self.recorder, on_start=self._on_start)

    def _on_start(self):
        self.recorder.start()
        self.overlay = RecordingOverlay(self.recorder, on_stop=self._on_stop)

    def _on_stop(self):
        if not self.recorder.steps:
            messagebox.showwarning("注意", "沒有記錄到任何步驟。")
            self.setup_win.win.deiconify()
            return
        self.editor = EditorWindow(self.recorder)

    def run(self):
        self.setup_win.run()


if __name__ == '__main__':
    app = App()
    app.run()
