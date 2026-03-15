from __future__ import annotations

import mss
import tkinter as tk


class ScreenOverlayBase(tk.Toplevel):
    def __init__(self, master: tk.Tk, title: str) -> None:
        super().__init__(master)
        with mss.mss() as sct:
            monitor = sct.monitors[0]

        self.offset_left = monitor["left"]
        self.offset_top = monitor["top"]
        self.geometry(
            f"{monitor['width']}x{monitor['height']}+{self.offset_left}+{self.offset_top}"
        )
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.28)
        self.configure(bg="#111111")

        self.canvas = tk.Canvas(self, bg="#111111", highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_text(
            24,
            24,
            anchor="nw",
            fill="#ffffff",
            font=("Microsoft JhengHei UI", 14, "bold"),
            text=title,
        )

        self.bind("<Escape>", lambda _: self.destroy())
        self.focus_force()
        self.grab_set()


class PointPicker(ScreenOverlayBase):
    def __init__(self, master: tk.Tk, title: str, on_pick) -> None:
        super().__init__(master, title)
        self.on_pick = on_pick
        self.canvas.bind("<Button-1>", self.pick)

    def pick(self, event) -> None:
        x = int(event.x_root)
        y = int(event.y_root)
        self.destroy()
        self.on_pick(x, y)


class RegionPicker(ScreenOverlayBase):
    def __init__(self, master: tk.Tk, title: str, on_pick) -> None:
        super().__init__(master, title)
        self.on_pick = on_pick
        self.start_x = 0
        self.start_y = 0
        self.preview_id = None
        self.canvas.bind("<ButtonPress-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.update_drag)
        self.canvas.bind("<ButtonRelease-1>", self.finish_drag)

    def start_drag(self, event) -> None:
        self.start_x = event.x_root
        self.start_y = event.y_root
        if self.preview_id is not None:
            self.canvas.delete(self.preview_id)
        self.preview_id = self.canvas.create_rectangle(
            event.x,
            event.y,
            event.x,
            event.y,
            outline="#20d5ff",
            width=2,
            dash=(4, 4),
        )

    def update_drag(self, event) -> None:
        if self.preview_id is None:
            return
        self.canvas.coords(
            self.preview_id,
            self.start_x - self.offset_left,
            self.start_y - self.offset_top,
            event.x_root - self.offset_left,
            event.y_root - self.offset_top,
        )

    def finish_drag(self, event) -> None:
        left = min(self.start_x, event.x_root)
        top = min(self.start_y, event.y_root)
        width = abs(event.x_root - self.start_x)
        height = abs(event.y_root - self.start_y)
        self.destroy()
        if width < 5 or height < 5:
            return
        self.on_pick(left, top, width, height)
