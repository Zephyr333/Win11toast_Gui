import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import asyncio
import os
import warnings
import json
import base64
from threading import Thread
import sys  # 新增
from winsdk.windows.data.xml.dom import XmlDocument
from winsdk.windows.foundation import IPropertyValue
from winsdk.windows.ui.notifications import (
    ToastNotificationManager,
    ToastNotification,
    NotificationData,
    ToastActivatedEventArgs
)

# Suppress libpng warnings
warnings.filterwarnings(
    "ignore", message=".*iCCP: known incorrect sRGB profile.*")

DEFAULT_APP_ID = 'Python'
CONFIG_FILE = 'config.json'

xml = """
<toast activationType="protocol" launch="http:" scenario="default">
    <visual>
        <binding template='ToastGeneric'></binding>
    </visual>
</toast>
"""


####################################################################
# 原有库函数（完整保留），针对需求添加图标形状、持续时间等
####################################################################

def set_attribute(document, xpath, name, value):
    attribute = document.create_attribute(name)
    attribute.value = value
    document.select_single_node(xpath).attributes.set_named_item(attribute)


def add_text(msg, document):
    if isinstance(msg, str):
        msg = {'text': msg}
    binding = document.select_single_node('//binding')
    text = document.create_element('text')
    for name, value in msg.items():
        if name == 'text':
            text.inner_text = msg['text']
        else:
            text.set_attribute(name, value)
    binding.append_child(text)


def add_icon(icon, document, shape='圆形'):  # 默认形状改为圆形
    if isinstance(icon, str):
        # shape决定hint-crop
        if shape == '圆形':
            hintcrop = 'circle'
        else:
            hintcrop = 'none'
        icon = {
            'placement': 'appLogoOverride',
            'hint-crop': hintcrop,
            'src': icon
        }
    binding = document.select_single_node('//binding')
    image = document.create_element('image')
    for name, value in icon.items():
        image.set_attribute(name, value)
    binding.append_child(image)


def add_image(img, document, placement='正文下方'):
    if isinstance(img, str):
        if placement == "标题上方":
            img = {'placement': 'hero', 'src': img}
        else:
            img = {'src': img}
    binding = document.select_single_node('//binding')
    image = document.create_element('image')
    for name, value in img.items():
        image.set_attribute(name, value)
    binding.append_child(image)


def add_progress(prog, document):
    binding = document.select_single_node('//binding')
    progress = document.create_element('progress')
    for name in prog:
        progress.set_attribute(name, '{' + name + '}')
    binding.append_child(progress)


def add_audio(aud, document, silent=False, scenario='default'):
    if isinstance(aud, dict):
        aud_attrs = aud
    elif isinstance(aud, str):
        aud_attrs = {'src': aud}
    else:
        aud_attrs = {}
    if silent:
        aud_attrs['silent'] = 'true'
    document_root = document.select_single_node('/toast')
    audio = document.create_element('audio')
    for name, value in aud_attrs.items():
        audio.set_attribute(name, value)
    if scenario != 'default':
        set_attribute(document, '/toast', 'scenario', scenario)
    document_root.append_child(audio)


def create_actions(document):
    toast = document.select_single_node('/toast')
    actions = document.create_element('actions')
    toast.append_child(actions)
    return actions


def add_button(button, document):
    if isinstance(button, str):
        button = {
            'activationType': 'protocol',
            'arguments': 'http:' + button,
            'content': button
        }
    actions = document.select_single_node(
        '//actions') or create_actions(document)
    action = document.create_element('action')
    for name, value in button.items():
        action.set_attribute(name, value)
    actions.append_child(action)


def add_input(id, document):
    if isinstance(id, str):
        id = {
            'id': id,
            'type': 'text',
            'placeHolderContent': id
        }
    actions = document.select_single_node(
        '//actions') or create_actions(document)
    input_elem = document.create_element('input')
    for name, value in id.items():
        input_elem.set_attribute(name, value)
    actions.append_child(input_elem)


def add_selection(selection, document):
    if isinstance(selection, list):
        selection_elements = selection
    else:
        selection_elements = [selection]
    actions = document.select_single_node(
        '//actions') or create_actions(document)
    input_elem = document.create_element('input')
    input_elem.set_attribute('id', 'selection')
    input_elem.set_attribute('type', 'selection')
    actions.append_child(input_elem)
    for sel in selection_elements:
        selection_element = document.create_element('selection')
        selection_element.set_attribute('id', sel)
        selection_element.set_attribute('content', sel)
        input_elem.append_child(selection_element)


result = list()


def result_wrapper(*args):
    global result
    result = args
    return result


def activated_args(_, event):
    global result
    e = ToastActivatedEventArgs._from(event)
    user_input = dict([
        (name, IPropertyValue._from(e.user_input[name]).get_string())
        for name in e.user_input
    ])
    result = {
        'arguments': e.arguments,
        'user_input': user_input
    }
    return result


async def play_sound(audio):
    from winsdk.windows.media.core import MediaSource
    from winsdk.windows.media.playback import MediaPlayer
    if not audio:
        return
    base = os.path.basename(audio)
    # 修改：不允许小写a、小写b和0~7作为首字母
    if not base or (base[0] in ('a', 'b') or (base[0].isdigit() and base[0] in '01234567')):
        messagebox.showerror("错误", "音频文件名首字母不能是小写a、小写b或数字0-7")
        return
    try:
        if audio.startswith('http'):
            from winsdk.windows.foundation import Uri
            source = MediaSource.create_from_uri(Uri(audio))
        else:
            from winsdk.windows.storage import StorageFile
            file = await StorageFile.get_file_from_path_async(audio)
            source = MediaSource.create_from_storage_file(file)
        player = MediaPlayer()
        player.source = source
        player.play()
        await asyncio.sleep(5)
    except Exception as e:
        messagebox.showerror("错误", f"音频播放错误: {e}")


async def speak(text):
    from winsdk.windows.media.core import MediaSource
    from winsdk.windows.media.playback import MediaPlayer
    from winsdk.windows.media.speechsynthesis import SpeechSynthesizer
    if not text:
        return
    try:
        synth = SpeechSynthesizer()
        stream = await synth.synthesize_text_to_stream_async(text)
        player = MediaPlayer()
        player.source = MediaSource.create_from_stream(
            stream, stream.content_type)
        player.play()
        await asyncio.sleep(5)
    except Exception as e:
        messagebox.showerror("错误", f"朗读文本错误: {e}")


async def recognize(ocr):
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.graphics.imaging import BitmapDecoder
    if isinstance(ocr, str):
        ocr = {'ocr': ocr}
    if ocr['ocr'].startswith('http'):
        from winsdk.windows.foundation import Uri
        from winsdk.windows.storage.streams import RandomAccessStreamReference
        ref = RandomAccessStreamReference.create_from_uri(Uri(ocr['ocr']))
        stream = await ref.open_read_async()
    else:
        from winsdk.windows.storage import StorageFile, FileAccessMode
        file = await StorageFile.get_file_from_path_async(ocr['ocr'])
        stream = await file.open_async(FileAccessMode.READ)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    if 'lang' in ocr:
        from winsdk.windows.globalization import Language
        if OcrEngine.is_language_supported(Language(ocr['lang'])):
            engine = OcrEngine.try_create_from_language(Language(ocr['lang']))
        else:
            class UnsupportedOcrResult:
                def __init__(self):
                    self.text = '请安装对应OCR语言包'
            return UnsupportedOcrResult()
    else:
        engine = OcrEngine.try_create_from_user_profile_languages()
    return await engine.recognize_async(bitmap)


def available_recognizer_languages():
    from winsdk.windows.media.ocr import OcrEngine
    for language in OcrEngine.get_available_recognizer_languages():
        print(language.display_name, language.language_tag)
    print('以管理员方式运行可安装OCR语言包')


def notify(title=None, body=None, on_click=None, icon=None, image=None,
           progress=None, audio=None, dialogue=None, duration=None,
           input=None, inputs=[], selection=None,
           button=None, buttons=[], xml=xml, app_id=DEFAULT_APP_ID,
           scenario='default', icon_shape='圆形', silent=False):
    document = XmlDocument()
    document.load_xml(xml)
    if on_click:
        set_attribute(document, '/toast', 'launch', on_click)
    if duration:
        set_attribute(document, '/toast', 'duration', duration)
    if title:
        add_text(title, document)
    if body:
        add_text(body, document)
    if input:
        add_input(input, document)
    if inputs:
        for inp in inputs:
            add_input(inp, document)
    if selection:
        add_selection(selection, document)
    if button:
        add_button(button, document)
    if buttons:
        for b in buttons:
            add_button(b, document)
    if icon:
        add_icon(icon, document, shape=icon_shape)
    if image:
        add_image(image, document)
    if progress:
        add_progress(progress, document)
    # 铃声
    if audio and not silent:
        add_audio({'src': audio}, document, silent=False, scenario=scenario)
    else:
        if audio:
            add_audio({'silent': 'true'}, document,
                      silent=True, scenario=scenario)
    if dialogue:
        # Dialogue is handled separately; ensure toast has title and body
        pass  # No need to add extra audio here

    notification = ToastNotification(document)
    if progress:
        data = NotificationData()
        for name, value in progress.items():
            data.values[name] = str(value)
        data.sequence_number = 1
        notification.data = data
        notification.tag = 'my_tag'

    if app_id == DEFAULT_APP_ID:
        try:
            notifier = ToastNotificationManager.create_toast_notifier()
        except:
            notifier = ToastNotificationManager.create_toast_notifier(app_id)
    else:
        notifier = ToastNotificationManager.create_toast_notifier(app_id)

    notifier.show(notification)
    return notification


async def toast_async(title=None, body=None, on_click=None, icon=None,
                      image=None, progress=None, audio=None, dialogue=None,
                      duration=None, input=None, inputs=[], selection=None,
                      button=None, buttons=[], xml=xml,
                      app_id=DEFAULT_APP_ID, ocr=None, on_dismissed=print,
                      on_failed=print, scenario='default', icon_shape='圆形',
                      silent=False):
    if ocr:
        title = 'OCR Result'
        body = (await recognize(ocr)).text
        src = ocr if isinstance(ocr, str) else ocr['ocr']
        image = {'placement': 'hero', 'src': src}

    # 发送通知
    notify(
        title=title,
        body=body,
        on_click=on_click,
        icon=icon,
        image=image,
        progress=progress,
        audio=audio,
        dialogue=dialogue,
        duration=duration,
        input=input,
        inputs=inputs,
        selection=selection,  # 修改这里
        button=button,
        buttons=buttons,
        xml=xml,
        app_id=app_id,
        scenario=scenario,
        icon_shape=icon_shape,
        silent=silent
    )


def toast(*args, **kwargs):
    asyncio.run(toast_async(**kwargs))


def update_progress(progress, app_id=DEFAULT_APP_ID):
    data = NotificationData()
    for name, value in progress.items():
        data.values[name] = str(value)
    data.sequence_number = 2
    if app_id == DEFAULT_APP_ID:
        try:
            notifier = ToastNotificationManager.create_toast_notifier()
        except:
            notifier = ToastNotificationManager.create_toast_notifier(app_id)
    else:
        notifier = ToastNotificationManager.create_toast_notifier(app_id)
    return notifier.update(data, 'my_tag')


####################################################################
# 以下是满足最新需求的大型GUI，包括输入框/选择框/按钮单独界面、图标形状、持续时间等
####################################################################

class LargeToastGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Win11自定义通知")
        self.geometry("950x500")  # 增加高度以适应新增控件

        # 缩放比例
        self.zoom_level = 1.0

        # 定义字体
        self.custom_font = ("楷体", 12)  # 设置为楷体，字号12

        # 使用 PanedWindow 以支持分隔条拖动
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # 左侧导航
        self.nav_frame = tk.Frame(self.paned, bg="#f0f0f0", width=200)
        self.paned.add(self.nav_frame, weight=1)

        # 右侧内容框架
        self.content_frame = tk.Frame(self.paned)
        self.paned.add(self.content_frame, weight=4)

        # 使用 Canvas 添加滚动功能
        self.canvas = tk.Canvas(self.content_frame)
        self.scrollbar_v = ttk.Scrollbar(
            self.content_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_h = ttk.Scrollbar(
            self.content_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(
            yscrollcommand=self.scrollbar_v.set, xscrollcommand=self.scrollbar_h.set)

        # 使用 grid 统一管理布局
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar_v.grid(row=0, column=1, sticky="ns")
        self.scrollbar_h.grid(row=1, column=0, sticky="ew")

        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # 创建一个框架在 Canvas 中
        self.main_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.main_frame, anchor='nw')

        # 绑定配置事件以更新滚动区域
        self.main_frame.bind("<Configure>", self.on_frame_configure)

        # 定义导航按钮
        nav_buttons = [
            ("基础设置", self.show_basic_frame),
            ("高级设置", self.show_advanced_frame),
            ("输入框管理", self.show_input_frame),
            ("选择框管理", self.show_selection_frame),
            ("按钮管理", self.show_button_frame),
            ("创建发送通知脚本", self.create_send_script),  # 新增按钮
            ("打开输出目录", self.open_output_directory)  # 新增按钮
        ]
        for (text, command) in nav_buttons:
            btn = tk.Button(self.nav_frame, text=text,
                            command=command, font=self.custom_font, height=2)
            btn.pack(fill=tk.X, pady=10, padx=10)  # 增加间距和按钮高度

        # ===== 基础设置Frame =====
        self.basic_frame = tk.Frame(self.main_frame)
        self.create_basic_frame()

        # ===== 高级设置Frame =====
        self.advanced_frame = tk.Frame(self.main_frame)
        self.create_advanced_frame()

        # ===== 输入框管理Frame =====
        self.input_frame = tk.Frame(self.main_frame)
        self.create_input_frame()

        # ===== 选择框管理Frame =====
        self.selection_frame = tk.Frame(self.main_frame)
        self.create_selection_frame()

        # ===== 按钮管理Frame =====
        self.button_frame = tk.Frame(self.main_frame)
        self.create_button_frame()

        # ===== 底部按钮（发送通知、重置、确定、取消）=====
        self.bottom_frame = tk.Frame(self.main_frame)
        self.bottom_frame.grid(
            row=100, column=0, columnspan=2, pady=20, sticky="ew")
        self.bottom_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        tk.Button(self.bottom_frame, text="发送通知", command=self.send_notification, font=self.custom_font, height=2).grid(
            row=0, column=0, padx=10, sticky="ew")
        tk.Button(self.bottom_frame, text="重置", command=self.reset_form, font=self.custom_font, height=2).grid(
            row=0, column=1, padx=10, sticky="ew")
        tk.Button(self.bottom_frame, text="确定", command=self.save_and_close, font=self.custom_font, height=2).grid(
            row=0, column=2, padx=10, sticky="ew")
        tk.Button(self.bottom_frame, text="取消", command=self.cancel_changes, font=self.custom_font, height=2).grid(
            row=0, column=3, padx=10, sticky="ew")

        # 记录初始状态
        self.old_state = self.load_config()
        self.load_state(self.old_state)
        # 默认显示基础设置
        self.show_basic_frame()

        # 绑定鼠标滚轮事件
        self.bind_events()

    def create_basic_frame(self):
        row_b = 0
        tk.Label(self.basic_frame, text="标题:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.title_var = tk.StringVar()
        tk.Entry(self.basic_frame, textvariable=self.title_var,
                 width=50, font=self.custom_font).grid(row=row_b, column=1, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="正文:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.body_var = tk.StringVar()
        tk.Entry(self.basic_frame, textvariable=self.body_var,
                 width=50, font=self.custom_font).grid(row=row_b, column=1, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="应用ID:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.app_id_var = tk.StringVar(value=DEFAULT_APP_ID)
        tk.Entry(self.basic_frame, textvariable=self.app_id_var,
                 width=50, font=self.custom_font).grid(row=row_b, column=1, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="图标路径:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.icon_var = tk.StringVar()
        tk.Entry(self.basic_frame, textvariable=self.icon_var,
                 width=50, font=self.custom_font).grid(row=row_b, column=1, padx=10, pady=10)
        tk.Button(self.basic_frame, text="浏览", command=self.select_icon, font=self.custom_font, width=10).grid(
            row=row_b, column=2, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="图标形状:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.icon_shape_var = tk.StringVar(value="圆形")  # 默认改为圆形
        shape_combo = ttk.Combobox(
            self.basic_frame, textvariable=self.icon_shape_var, values=["圆形", "方形"], width=48, font=self.custom_font)
        shape_combo.grid(row=row_b, column=1, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="图片路径:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.image_var = tk.StringVar()
        tk.Entry(self.basic_frame, textvariable=self.image_var,
                 width=50, font=self.custom_font).grid(row=row_b, column=1, padx=10, pady=10)
        tk.Button(self.basic_frame, text="浏览", command=self.select_image, font=self.custom_font, width=10).grid(
            row=row_b, column=2, padx=10, pady=10)
        row_b += 1

        tk.Label(self.basic_frame, text="图片位置:", font=self.custom_font).grid(
            row=row_b, column=0, sticky="w", padx=10, pady=10)
        self.image_placement_var = tk.StringVar(value="正文下方")
        image_combo = ttk.Combobox(self.basic_frame, textvariable=self.image_placement_var, values=[
                                   "标题上方", "正文下方"], width=48, font=self.custom_font)
        image_combo.grid(row=row_b, column=1, padx=10, pady=10)
        row_b += 1

    def create_advanced_frame(self):
        row_a = 0

        tk.Label(self.advanced_frame, text="铃声类型:", font=self.custom_font).grid(
            row=row_a, column=0, sticky="w", padx=10, pady=10)
        self.ring_var = tk.StringVar()
        ring_combo = ttk.Combobox(
            self.advanced_frame,
            textvariable=self.ring_var,
            values=[
                "默认",
                "提醒",
                "来电",
                "邮件",
                "短信"
            ],
            width=48,
            font=self.custom_font
        )
        ring_combo.grid(row=row_a, column=1, padx=10, pady=10)
        ring_combo.set("默认")
        row_a += 1

        # 铃声类型值映射
        self.ring_mapping = {
            "默认": "ms-winsoundevent:Notification.Default",
            "提醒": "ms-winsoundevent:Notification.Reminder",
            "来电": "ms-winsoundevent:Notification.IncomingCall",
            "邮件": "ms-winsoundevent:Notification.Mail",
            "短信": "ms-winsoundevent:Notification.SMS"
        }

        tk.Label(self.advanced_frame, text="持续时间:", font=self.custom_font).grid(
            row=row_a, column=0, sticky="w", padx=10, pady=10)
        self.duration_var = tk.StringVar()
        duration_combo = ttk.Combobox(
            self.advanced_frame,
            textvariable=self.duration_var,
            values=["短 (7s)", "长 (25s)"],
            width=48,
            font=self.custom_font
        )
        duration_combo.grid(row=row_a, column=1, padx=10, pady=10)
        duration_combo.set("短 (7s)")
        row_a += 1

        tk.Label(self.advanced_frame, text="音频路径:", font=self.custom_font).grid(
            row=row_a, column=0, sticky="w", padx=10, pady=10)
        self.audio_var = tk.StringVar()
        tk.Entry(self.advanced_frame, textvariable=self.audio_var,
                 width=50, font=self.custom_font).grid(row=row_a, column=1, padx=10, pady=10)
        tk.Button(self.advanced_frame, text="浏览", command=self.select_audio, font=self.custom_font, width=10).grid(
            row=row_a, column=2, padx=10, pady=10)
        row_a += 1

        # 音频播放选项（仅静音在一行）
        options_frame = tk.Frame(self.advanced_frame)
        options_frame.grid(row=row_a, column=1, sticky="w", padx=10, pady=10)
        self.silent_var = tk.BooleanVar()
        tk.Checkbutton(options_frame, text="静音", variable=self.silent_var, font=self.custom_font).pack(
            side=tk.LEFT, padx=(0, 10))
        row_a += 1

        tk.Label(self.advanced_frame, text="朗读文本:", font=self.custom_font).grid(
            row=row_a, column=0, sticky="w", padx=10, pady=10)
        self.speech_var = tk.StringVar()
        tk.Entry(self.advanced_frame, textvariable=self.speech_var,
                 width=50, font=self.custom_font).grid(row=row_a, column=1, padx=10, pady=10)
        row_a += 1

        # 点击后跳转URL
        tk.Label(self.advanced_frame, text="点击跳转:", font=self.custom_font).grid(
            row=row_a, column=0, sticky="w", padx=10, pady=10)
        self.url_var = tk.StringVar()
        tk.Entry(self.advanced_frame, textvariable=self.url_var,
                 width=50, font=self.custom_font).grid(row=row_a, column=1, padx=10, pady=10)
        row_a += 1

    def create_input_frame(self):
        self.input_list = []
        self.input_listbox = tk.Listbox(
            self.input_frame, width=60, height=15, font=self.custom_font)
        self.input_listbox.grid(
            row=0, column=0, columnspan=3, padx=20, pady=20)
        tk.Button(self.input_frame, text="添加输入框", command=self.add_input_item, font=self.custom_font, width=15).grid(
            row=1, column=0, sticky="w", padx=20, pady=10)
        tk.Button(self.input_frame, text="删除选中输入框", command=self.remove_input_item, font=self.custom_font, width=20).grid(
            row=1, column=1, sticky="w", padx=10, pady=10)

    def create_selection_frame(self):
        self.selection_list = []
        self.selection_listbox = tk.Listbox(
            self.selection_frame, width=60, height=15, font=self.custom_font)
        self.selection_listbox.grid(
            row=0, column=0, columnspan=3, padx=20, pady=20)
        tk.Button(self.selection_frame, text="添加选择项", command=self.add_selection_item, font=self.custom_font, width=15).grid(
            row=1, column=0, sticky="w", padx=20, pady=10)
        tk.Button(self.selection_frame, text="删除选中选择项", command=self.remove_selection_item, font=self.custom_font, width=20).grid(
            row=1, column=1, sticky="w", padx=10, pady=10)

    def create_button_frame(self):
        self.button_list = []
        self.button_listbox = tk.Listbox(
            self.button_frame, width=60, height=15, font=self.custom_font)
        self.button_listbox.grid(
            row=0, column=0, columnspan=3, padx=20, pady=20)
        tk.Button(self.button_frame, text="添加按钮", command=self.add_button_item, font=self.custom_font, width=15).grid(
            row=1, column=0, sticky="w", padx=20, pady=10)
        tk.Button(self.button_frame, text="删除选中按钮", command=self.remove_button_item, font=self.custom_font, width=20).grid(
            row=1, column=1, sticky="w", padx=10, pady=10)

    # ===== 导航切换 =====
    def show_basic_frame(self):
        self.hide_all_frames()
        self.basic_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_advanced_frame(self):
        self.hide_all_frames()
        self.advanced_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_input_frame(self):
        self.hide_all_frames()
        self.input_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_selection_frame(self):
        self.hide_all_frames()
        self.selection_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_button_frame(self):
        self.hide_all_frames()
        self.button_frame.grid(
            row=0, column=0, sticky="nsew", padx=20, pady=20)

    def hide_all_frames(self):
        frames = [self.basic_frame, self.advanced_frame,
                  self.input_frame, self.selection_frame, self.button_frame]
        for frame in frames:
            frame.grid_forget()

    # ===== 选择文件 =====
    def select_icon(self):
        path = filedialog.askopenfilename()
        if path:
            self.icon_var.set(path.replace('/', '\\'))  # 替换为反斜杠

    def select_image(self):
        path = filedialog.askopenfilename()
        if path:
            self.image_var.set(path.replace('/', '\\'))  # 替换为反斜杠

    def select_audio(self):
        path = filedialog.askopenfilename()
        if path:
            path = path.replace('/', '\\')  # 替换为反斜杠
            base = os.path.basename(path)
            # 修改：不允许小写a、小写b和0~7开头
            if not base or (base[0] in ('a', 'b') or (base[0].isdigit() and base[0] in '01234567')):
                # 显示错误弹窗说明不允许的首字母
                messagebox.showerror("错误", "音频文件名首字母不能是小写a、小写b或数字0-7")
                self.audio_var.set("")
            else:
                self.audio_var.set(path)

    # ===== 输入框管理 =====
    def add_input_item(self):
        name = simpledialog.askstring(
            "输入框名称", "请输入输入框名称:")
        if name:
            new_item = name
            self.input_list.append(new_item)
            self.input_listbox.insert(tk.END, new_item)

    def remove_input_item(self):
        sel = self.input_listbox.curselection()
        if not sel:
            return
        index = sel[0]
        self.input_listbox.delete(index)
        del self.input_list[index]

    # ===== 选择框管理 =====
    def add_selection_item(self):
        name = simpledialog.askstring(
            "选择项名称", "请输入选择项名称:")
        if name:
            new_item = name
            self.selection_list.append(new_item)
            self.selection_listbox.insert(tk.END, new_item)

    def remove_selection_item(self):
        sel = self.selection_listbox.curselection()
        if not sel:
            return
        index = sel[0]
        self.selection_listbox.delete(index)
        del self.selection_list[index]

    # ===== 按钮管理 =====
    def add_button_item(self):
        name = simpledialog.askstring(
            "按钮名称", "请输入按钮名称:")
        if name:
            new_item = name
            self.button_list.append(new_item)
            self.button_listbox.insert(tk.END, new_item)

    def remove_button_item(self):
        sel = self.button_listbox.curselection()
        if not sel:
            return
        index = sel[0]
        self.button_listbox.delete(index)
        del self.button_list[index]

    # ===== 发送通知 =====
    def send_notification(self):
        st = self.get_current_state()

        # 图片位置
        image_data = None
        if st["image"]:
            if st["image_placement"] == "标题上方":
                image_data = {"placement": "hero", "src": st["image"]}
            else:
                image_data = {"src": st["image"]}

        # 铃声类型值
        ring_value = self.ring_mapping.get(
            st["ring"], "ms-winsoundevent:Notification.Default")

        # 持续时间
        duration_mapping = {
            "短 (7s)": "short",
            "长 (25s)": "long"
        }
        duration = duration_mapping.get(st["duration"], "short")

        # 点击后跳转URL
        url = st.get("url", "")

        # 音频播放选项
        silent = st["silent"]

        # 获取输入框、选择框、按钮
        inputs = st["input_list"]
        selection = st["selection_list"]  # 修改这里
        buttons = st["button_list"]

        def send():
            # 为每个线程创建独立的事件循环
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(self.async_send_notification(
                    st, image_data, ring_value, duration, url, silent, inputs, selection, buttons))
            finally:
                loop.close()

        Thread(target=send).start()

    async def async_send_notification(self, st, image_data, ring_value, duration, url, silent, inputs, selection, buttons):
        try:
            await toast_async(
                title=st["title"],
                body=st["body"],
                app_id=st["app_id"] or DEFAULT_APP_ID,
                icon=st["icon"] or None,
                icon_shape=st["icon_shape"],
                image=image_data,
                audio=ring_value if ring_value else None,
                dialogue=st["speech"] if st["speech"] else None,
                duration=duration,
                on_click=url if url else None,
                silent=silent,
                scenario="IncomingCall" if ring_value == "ms-winsoundevent:Notification.IncomingCall" else "default",
                inputs=inputs,
                selection=selection,  # 修改这里
                buttons=buttons
            )

            # 如果有朗读文本，执行朗读
            if st["speech"]:
                await speak(st["speech"])

            # 播放自定义音频
            if st["audio"] and not silent:
                await play_sound(st["audio"])
        except Exception as e:
            # 在发送通知失败时弹出错误消息
            messagebox.showerror("错误", f"发送通知失败: {e}")

    # ===== 重置、确定、取消 =====
    def reset_form(self):
        self.title_var.set("")
        self.body_var.set("")
        self.app_id_var.set(DEFAULT_APP_ID)
        self.icon_var.set("")
        self.icon_shape_var.set("圆形")
        self.image_var.set("")
        self.image_placement_var.set("正文下方")
        self.ring_var.set("默认")
        self.duration_var.set("短 (7s)")
        self.audio_var.set("")
        self.speech_var.set("")
        self.url_var.set("")
        self.silent_var.set(False)
        self.input_list.clear()
        self.input_listbox.delete(0, tk.END)
        self.selection_list.clear()
        self.selection_listbox.delete(0, tk.END)
        self.button_list.clear()
        self.button_listbox.delete(0, tk.END)

    def save_and_close(self):
        self.save_config(self.get_current_state())
        self.destroy()

    def cancel_changes(self):
        self.load_state(self.old_state)
        self.destroy()

    # ===== 状态存取 =====
    def get_current_state(self):
        return {
            "title": self.title_var.get(),
            "body": self.body_var.get(),
            "app_id": self.app_id_var.get(),
            "icon": self.icon_var.get(),
            "icon_shape": self.icon_shape_var.get(),
            "image": self.image_var.get(),
            "image_placement": self.image_placement_var.get(),
            "ring": self.ring_var.get(),
            "duration": self.duration_var.get(),
            "audio": self.audio_var.get(),
            "speech": self.speech_var.get(),
            "url": self.url_var.get(),
            "silent": self.silent_var.get(),
            "input_list": list(self.input_list),
            "selection_list": list(self.selection_list),
            "button_list": list(self.button_list)
        }

    def load_state(self, st):
        self.title_var.set(st["title"])
        self.body_var.set(st["body"])
        self.app_id_var.set(st["app_id"])
        self.icon_var.set(st["icon"])
        self.icon_shape_var.set("圆形" if st["icon_shape"] not in [
                                "圆形", "方形"] else st["icon_shape"])
        self.image_var.set(st["image"])
        self.image_placement_var.set(st["image_placement"])
        self.ring_var.set(st["ring"])
        self.duration_var.set(st["duration"])
        self.audio_var.set(st["audio"])
        self.speech_var.set(st["speech"])
        self.url_var.set(st["url"])
        self.silent_var.set(st["silent"])

        self.input_list = list(st["input_list"])
        self.input_listbox.delete(0, tk.END)
        for i in self.input_list:
            self.input_listbox.insert(tk.END, i)

        self.selection_list = list(st["selection_list"])
        self.selection_listbox.delete(0, tk.END)
        for s in self.selection_list:
            self.selection_listbox.insert(tk.END, s)

        self.button_list = list(st["button_list"])
        self.button_listbox.delete(0, tk.END)
        for b in self.button_list:
            self.button_listbox.insert(tk.END, b)

    def load_config(self):
        base_path = self.get_base_path()  # 使用新的基路径
        config_path = os.path.join(base_path, CONFIG_FILE)
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return self.get_current_state()

    def save_config(self, config):
        base_path = self.get_base_path()  # 使用新的基路径
        config_path = os.path.join(base_path, CONFIG_FILE)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)

    # ===== 添加滚动条和绑定事件 =====
    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def bind_events(self):
        # 绑定鼠标滚轮事件
        self.bind_all("<MouseWheel>", self.on_mousewheel)
        self.bind_all("<Shift-MouseWheel>", self.on_shift_mousewheel)
        self.bind_all("<Control-MouseWheel>", self.on_ctrl_mousewheel)

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")

    def on_ctrl_mousewheel(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    def zoom_in(self):
        self.zoom_level *= 1.1
        self.apply_zoom()

    def zoom_out(self):
        self.zoom_level /= 1.1
        self.apply_zoom()

    def apply_zoom(self):
        # Tkinter Frame does not support scaling directly.
        # Instead, adjust font sizes as a workaround.
        # Prevent font size from being too small
        new_size = max(int(12 * self.zoom_level), 10)  # 基础字号12
        self.scale_widget(self.main_frame, new_size)

    def scale_widget(self, widget, size):
        if isinstance(widget, (tk.Label, tk.Button, ttk.Combobox, tk.Entry, tk.Listbox)):
            try:
                widget.configure(font=(self.custom_font[0], size))
            except:
                pass
        for child in widget.winfo_children():
            self.scale_widget(child, size)

    # ===== 创建发送通知脚本 =====
    def create_send_script(self):
        script_name = simpledialog.askstring(
            "创建发送通知脚本", "请输入脚本名称（无需加 .py 后缀）:")
        if not script_name:
            return
        script_name = script_name.strip()
        if not script_name:
            messagebox.showerror("错误", "脚本名称不能为空。")
            return
        base_path = self.get_base_path()  # 获取基路径
        script_filename = f"{script_name}.py"
        script_path = os.path.join(base_path, script_filename)  # 使用基路径
        if os.path.exists(script_path):
            if not messagebox.askyesno("确认", f"{script_filename} 已存在，是否覆盖？"):
                return
        try:
            # 获取当前状态
            st = self.get_current_state()

            # 图片位置
            if st["image"]:
                if st["image_placement"] == "标题上方":
                    image_data = {"placement": "hero", "src": st["image"]}
                else:
                    image_data = {"src": st["image"]}
            else:
                image_data = None

            # 铃声类型值
            ring_value = self.ring_mapping.get(
                st["ring"], "ms-winsoundevent:Notification.Default")

            # 持续时间
            duration_mapping = {
                "短 (7s)": "short",
                "长 (25s)": "long"
            }
            duration = duration_mapping.get(st["duration"], "short")

            # 点击后跳转URL
            url = st.get("url", "")

            # 音频播放选项
            silent = st["silent"]

            # 获取输入框、选择框、按钮
            inputs = st["input_list"]
            selection = st["selection_list"]
            buttons = st["button_list"]

            # 读取并编码图标、图片和音频文件
            icon_b64 = ""
            if st["icon"]:
                with open(st["icon"], "rb") as img_file:
                    icon_b64 = base64.b64encode(
                        img_file.read()).decode('utf-8')

            image_b64 = ""
            if st["image"]:
                with open(st["image"], "rb") as img_file:
                    image_b64 = base64.b64encode(
                        img_file.read()).decode('utf-8')

            audio_b64 = ""
            if st["audio"]:
                with open(st["audio"], "rb") as audio_file:
                    audio_b64 = base64.b64encode(
                        audio_file.read()).decode('utf-8')

            # 构造脚本内容
            # 修改 image 部分，基于 image_placement 设定 placement
            image_placement_condition = (
                '{"placement": "hero", "src": image_path}' if st["image_placement"] == "标题上方" else '{"src": image_path}') if st["image"] else "None"

            script_content = f"""# {script_filename}
import asyncio
import base64
import os
import tempfile
from winsdk.windows.ui.notifications import ToastNotificationManager
from Win11toast_Gui import toast_async, play_sound, speak, DEFAULT_APP_ID

def decode_and_save(data_b64, suffix):
    if not data_b64:
        return None
    decoded_data = base64.b64decode(data_b64)
    temp_dir = tempfile.gettempdir()  # 使用系统临时目录
    fd, path = tempfile.mkstemp(suffix=suffix, dir=temp_dir)
    os.close(fd)
    with open(path, 'wb') as f:
        f.write(decoded_data)
    return path

async def main():
    icon_path = decode_and_save("{icon_b64}", ".png") if "{icon_b64}" else None
    image_path = decode_and_save("{image_b64}", ".png") if "{image_b64}" else None
    audio_path = decode_and_save("{audio_b64}", ".mp3") if "{audio_b64}" else None

    await toast_async(
        title={json.dumps(st["title"])},
        body={json.dumps(st["body"])},
        app_id={json.dumps(st["app_id"])} or DEFAULT_APP_ID,
        icon=icon_path,
        icon_shape={json.dumps(st["icon_shape"])},
        image={image_placement_condition},
        audio={json.dumps(ring_value)} if "{ring_value}" else None,
        dialogue={json.dumps(st["speech"])} if "{st["speech"]}" else None,
        duration={json.dumps(st["duration"])},
        on_click={json.dumps(url)} if "{url}" else None,
        silent={str(silent)},
        scenario="{ 'IncomingCall' if ring_value == 'ms-winsoundevent:Notification.IncomingCall' else 'default' }",
        inputs={json.dumps(inputs)},
        selection={json.dumps(selection)},
        buttons={json.dumps(buttons)}
    )

    # 如果有朗读文本，执行朗读
    {"await speak(" + json.dumps(st["speech"]) + ")" if st["speech"] else ""}

    # 播放自定义音频
    {"await play_sound(audio_path)" if st["audio"] and not silent else ""}

    # 等待一段时间以确保通知加载完成
    await asyncio.sleep(2)

    # 删除临时文件
    if icon_path and os.path.exists(icon_path):
        os.remove(icon_path)
    if image_path and os.path.exists(image_path):
        os.remove(image_path)
    if audio_path and os.path.exists(audio_path):
        os.remove(audio_path)

if __name__ == "__main__":
    asyncio.run(main())
"""

            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            messagebox.showinfo("成功", f"脚本已创建在 {script_path}")
        except Exception as e:
            messagebox.showerror("错误", f"创建脚本失败: {e}")

    # ===== 打开输出目录 =====
    def open_output_directory(self):
        try:
            base_path = self.get_base_path()  # 获取基路径
            os.startfile(base_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开输出目录: {e}")

    # ===== 获取基路径 =====
    def get_base_path(self):
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe，获取exe所在的目录
            return os.path.dirname(sys.executable)
        else:
            # 如果是开发环境，获取脚本所在的目录
            return os.path.dirname(os.path.abspath(__file__))


if __name__ == "__main__":
    app = LargeToastGUI()
    app.mainloop()
