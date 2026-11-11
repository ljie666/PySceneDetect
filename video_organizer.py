"""Display metadata and shot boundaries for a local MP4 file."""

import csv
import json
from io import BytesIO
from pathlib import Path
import subprocess
from queue import Empty, Queue
from threading import Thread
from typing import Callable, TypeVar
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from PIL import Image, ImageDraw
from scenedetect import ContentDetector, FrameTimecode, detect

Shot = tuple[FrameTimecode, FrameTimecode]
Result = TypeVar("Result")
SCENE_COLUMNS = ["文件名", "镜头编号", "开始时间", "结束时间", "时长（秒）"]


def read_video(path: Path) -> tuple[float, int, int]:
    result = subprocess.run(
        [
            "ffprobe", "-v", "error", "-select_streams", "V:0",
            "-show_entries", "format=duration:stream=width,height",
            "-of", "json", str(path.resolve()),
        ],
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        check=True,
        timeout=15,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    metadata = json.loads(result.stdout)
    stream = metadata["streams"][0]
    return float(metadata["format"]["duration"]), stream["width"], stream["height"]


def save_contact_sheet(path: Path, shots: list[Shot], output: Path) -> None:
    tile_width, tile_height = 240, 184
    columns = min(4, len(shots))
    rows = (len(shots) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * tile_width, rows * tile_height), "#20242b")
    draw = ImageDraw.Draw(sheet)
    for index, (start, end) in enumerate(shots):
        midpoint = (start.get_seconds() + end.get_seconds()) / 2
        frame = subprocess.run(
            [
                "ffmpeg", "-v", "error", "-ss", f"{midpoint:.6f}", "-i", str(path.resolve()),
                "-map", "0:V:0", "-frames:v", "1",
                "-vf", "scale=224:140:force_original_aspect_ratio=decrease:reset_sar=1",
                "-c:v", "png", "-f", "image2pipe", "pipe:1",
            ],
            capture_output=True,
            check=True,
            timeout=30,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        x = (index % columns) * tile_width
        y = (index // columns) * tile_height
        with Image.open(BytesIO(frame.stdout)) as thumbnail:
            sheet.paste(thumbnail, (x + (tile_width - thumbnail.width) // 2, y + 8 + (140 - thumbnail.height) // 2))
        draw.text((x + 8, y + 155), f"{index + 1}   {start.get_timecode()}", fill="white")
    sheet.save(output, format="PNG")


def save_scene_csv(path: Path, shots: list[Shot], output: Path) -> None:
    filename = path.name
    if filename.lstrip().startswith(("=", "+", "-", "@")):
        filename = "'" + filename

    with output.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(SCENE_COLUMNS)
        for number, (start, end) in enumerate(shots, start=1):
            duration = end.get_seconds() - start.get_seconds()
            writer.writerow([filename, number, start.get_timecode(), end.get_timecode(), f"{duration:.3f}"])


def save_scene_xlsx(path: Path, shots: list[Shot], output: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "镜头时间表"
    sheet.append(SCENE_COLUMNS)
    sheet.freeze_panes = "A2"
    for column, width in zip("ABCDE", (30, 12, 18, 18, 14)):
        sheet.column_dimensions[column].width = width

    for number, (start, end) in enumerate(shots, start=1):
        duration = end.get_seconds() - start.get_seconds()
        sheet.append([path.name, number, start.get_timecode(), end.get_timecode(), round(duration, 3)])
        row = sheet.max_row
        for column in (1, 3, 4):
            cell = sheet.cell(row, column)
            cell.data_type = "s"
            cell.number_format = "@"
        sheet.cell(row, 5).number_format = "0.000"
    workbook.save(output)


def save_folder_inventory(folder: Path, output: Path) -> tuple[int, int]:
    videos = sorted(
        (path for path in folder.rglob("*") if path.suffix.lower() == ".mp4" and path.is_file()),
        key=lambda path: path.relative_to(folder).as_posix().casefold(),
    )
    if not videos:
        raise ValueError("这个文件夹及其子文件夹里没有 MP4。")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "视频素材清单"
    sheet.append(["相对路径", "时长（秒）", "宽度", "高度", "读取状态"])
    sheet.freeze_panes = "A2"
    for column, width in zip("ABCDE", (56, 14, 12, 12, 16)):
        sheet.column_dimensions[column].width = width

    failed = 0
    for video in videos:
        relative_path = video.relative_to(folder).as_posix()
        try:
            duration, width, height = read_video(video)
        except FileNotFoundError as error:
            raise RuntimeError("找不到 ffprobe，请检查 FFmpeg 是否可用。") from error
        except (OSError, subprocess.SubprocessError, ValueError, KeyError, IndexError):
            sheet.append([relative_path, None, None, None, "读取失败"])
            failed += 1
        else:
            sheet.append([relative_path, round(duration, 3), width, height, "已读取"])
            sheet.cell(sheet.max_row, 2).number_format = "0.000"
        sheet.cell(sheet.max_row, 1).data_type = "s"

    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(output)
    return len(videos) - failed, failed


class Organizer(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("视频素材整理器")
        self.minsize(600, 360)
        self.video_path: Path | None = None
        self.shots: list[Shot] = []
        self.scene_status = tk.StringVar()
        self.filename = tk.StringVar(value="—")
        self.duration = tk.StringVar(value="—")
        self.resolution = tk.StringVar(value="—")

        panel = ttk.Frame(self, padding=24)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(1, weight=1)
        panel.rowconfigure(5, weight=1)
        toolbar = ttk.Frame(panel)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 16))
        self.choose_button = ttk.Button(toolbar, text="选择 MP4", command=self.choose_video)
        self.choose_button.pack(side="left")
        self.detect_button = ttk.Button(
            toolbar, text="检测镜头", command=self.detect_scenes, state="disabled"
        )
        self.detect_button.pack(side="left", padx=(8, 0))
        self.sheet_button = ttk.Button(
            toolbar, text="保存缩略图拼图", command=self.export_sheet, state="disabled"
        )
        self.sheet_button.pack(side="left", padx=(8, 0))
        self.csv_button = ttk.Button(
            toolbar, text="导出 CSV",
            command=lambda: self.export_table("CSV", ".csv", save_scene_csv), state="disabled"
        )
        self.csv_button.pack(side="left", padx=(8, 0))
        self.xlsx_button = ttk.Button(
            toolbar, text="导出 Excel",
            command=lambda: self.export_table("Excel", ".xlsx", save_scene_xlsx), state="disabled"
        )
        self.xlsx_button.pack(side="left", padx=(8, 0))
        for row, (label, value) in enumerate(
            [("文件名", self.filename), ("时长", self.duration), ("分辨率", self.resolution)],
            start=1,
        ):
            ttk.Label(panel, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 20), pady=6)
            ttk.Label(panel, textvariable=value, wraplength=300).grid(
                row=row, column=1, sticky="nw", pady=6
            )

        status_bar = ttk.Frame(panel)
        status_bar.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 4))
        ttk.Label(status_bar, textvariable=self.scene_status).pack(side="left")
        self.folder_button = ttk.Button(status_bar, text="导出文件夹清单", command=self.export_folder)
        self.folder_button.pack(side="right")
        table = ttk.Frame(panel)
        table.grid(row=5, column=0, columnspan=2, sticky="nsew")
        self.scenes = ttk.Treeview(table, columns=("shot", "start", "end"), show="headings", height=6)
        for column, title, width in (
            ("shot", "镜头", 65), ("start", "开始时间", 175), ("end", "结束时间", 175)
        ):
            self.scenes.heading(column, text=title)
            self.scenes.column(column, width=width, anchor="center")
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.scenes.yview)
        self.scenes.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.scenes.pack(side="left", fill="both", expand=True)

    def choose_video(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self, title="选择视频", filetypes=[("MP4 视频", "*.mp4")]
        )
        if not selected:
            return

        path = Path(selected)
        try:
            duration, width, height = read_video(path)
        except FileNotFoundError:
            messagebox.showerror("读取失败", "找不到 ffprobe，请检查 FFmpeg 是否可用。", parent=self)
            return
        except (OSError, subprocess.SubprocessError, ValueError, KeyError, IndexError):
            messagebox.showerror("读取失败", "无法读取视频信息，请选择可正常播放的 MP4。", parent=self)
            return

        self.filename.set(path.name)
        self.duration.set(f"{duration:.2f} 秒")
        self.resolution.set(f"{width} × {height}")
        self.video_path = path
        self.shots = []
        self.sheet_button.state(["disabled"])
        self.csv_button.state(["disabled"])
        self.xlsx_button.state(["disabled"])
        self.scenes.delete(*self.scenes.get_children())
        self.scene_status.set("")
        self.detect_button.state(["!disabled"])

    def detect_scenes(self) -> None:
        path = self.video_path
        if path is None:
            return

        self.shots = []
        self.scenes.delete(*self.scenes.get_children())
        self.scene_status.set("正在检测镜头…")
        self.run_task(
            lambda: detect(path, ContentDetector(), start_in_scene=True),
            self.show_scenes,
            "镜头检测失败",
        )

    def export_sheet(self) -> None:
        path, shots = self.video_path, self.shots
        if path is None or not shots:
            return

        selected = filedialog.asksaveasfilename(
            parent=self, title="保存缩略图拼图", initialfile=f"{path.stem}-缩略图.png",
            defaultextension=".png", filetypes=[("PNG 图片", "*.png")],
        )
        if not selected:
            return
        output = Path(selected)
        if output.resolve() == path.resolve():
            messagebox.showerror("无法保存", "请选择与原视频不同的文件名。", parent=self)
            return
        self.scene_status.set("正在生成缩略图拼图…")
        self.run_task(
            lambda: save_contact_sheet(path, shots, output),
            lambda _: self.scene_status.set(f"已保存拼图 · {len(shots)} 个镜头"),
            "保存拼图失败",
        )

    def export_table(
        self, label: str, extension: str, save: Callable[[Path, list[Shot], Path], None]
    ) -> None:
        path, shots = self.video_path, self.shots
        if path is None or not shots:
            return

        selected = filedialog.asksaveasfilename(
            parent=self, title="导出镜头时间表", initialfile=f"{path.stem}-镜头{extension}",
            defaultextension=extension, filetypes=[(f"{label} 表格", f"*{extension}")],
        )
        if not selected:
            return
        output = Path(selected)
        if output.suffix.lower() != extension:
            messagebox.showerror("无法导出", f"请使用 {extension} 扩展名。", parent=self)
            return
        if output.resolve() == path.resolve() or (
            output.exists() and path.exists() and output.samefile(path)
        ):
            messagebox.showerror("无法导出", "请选择与原视频不同的文件名。", parent=self)
            return
        self.run_task(
            lambda: save(path, shots, output),
            lambda _: self.scene_status.set(f"已导出 {label} · {len(shots)} 个镜头"),
            f"导出 {label} 失败",
        )

    def export_folder(self) -> None:
        selected = filedialog.askdirectory(parent=self, title="选择素材文件夹", mustexist=True)
        if not selected:
            return
        folder = Path(selected)
        selected = filedialog.asksaveasfilename(
            parent=self, title="保存视频素材清单", initialfile=f"{folder.name or '视频'}-素材清单.xlsx",
            initialdir=folder, defaultextension=".xlsx", filetypes=[("Excel 表格", "*.xlsx")],
        )
        if not selected:
            return
        output = Path(selected)
        if output.suffix.lower() != ".xlsx" or output.resolve().suffix.lower() != ".xlsx":
            messagebox.showerror("无法导出", "请选择 .xlsx 文件。", parent=self)
            return

        self.scene_status.set("正在扫描文件夹…")
        self.run_task(
            lambda: save_folder_inventory(folder, output),
            lambda counts: self.scene_status.set(f"清单已保存 · 成功 {counts[0]}，失败 {counts[1]}"),
            "导出清单失败",
        )

    def run_task(
        self, action: Callable[[], Result], on_success: Callable[[Result], None], error_title: str
    ) -> None:
        buttons = (
            self.choose_button, self.detect_button, self.sheet_button,
            self.csv_button, self.xlsx_button, self.folder_button,
        )
        for button in buttons:
            button.state(["disabled"])
        results: Queue[Result | Exception] = Queue()

        def analyze() -> None:
            try:
                results.put(action())
            except Exception as error:
                results.put(error)

        def poll() -> None:
            try:
                result = results.get_nowait()
            except Empty:
                self.after(80, poll)
                return
            self.choose_button.state(["!disabled"])
            self.folder_button.state(["!disabled"])
            if self.video_path is not None:
                self.detect_button.state(["!disabled"])
            if isinstance(result, Exception):
                self.scene_status.set(error_title)
                messagebox.showerror(error_title, str(result), parent=self)
            else:
                on_success(result)
            if self.shots:
                self.sheet_button.state(["!disabled"])
                self.csv_button.state(["!disabled"])
                self.xlsx_button.state(["!disabled"])

        Thread(target=analyze, daemon=True).start()
        self.after(80, poll)

    def show_scenes(self, shots: list[Shot]) -> None:
        self.shots = shots
        for number, (start, end) in enumerate(shots, start=1):
            self.scenes.insert("", "end", values=(number, start.get_timecode(), end.get_timecode()))
        self.scene_status.set(f"共 {len(shots)} 个镜头")


if __name__ == "__main__":
    Organizer().mainloop()
