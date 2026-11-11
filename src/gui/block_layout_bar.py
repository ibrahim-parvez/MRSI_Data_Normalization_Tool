"""
The Normalization block arranger: one segmented control whose three segments
are the blocks drawn across the Normalization sheet. Click a segment to
include or drop it, drag it to move it; left to right is the order it lands in
on the sheet.

The segments are positioned by hand rather than by a layout, so a dragged one
can follow the mouse while the others slide out of its way and leave the gap
it will drop into.

Two rules hold the arrangement together:
  * at least one data block (Carbonate or CO2) is always included, since there
    is nothing to normalize otherwise;
  * Yield is never left-most - it has no reference materials box to put in the
    sheet's top-left corner - and never leads the included blocks either, so
    the control always shows exactly what the sheet will do.
"""

from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QFrame, QLabel, QGraphicsDropShadowEffect
)
from PyQt6.QtCore import (
    Qt, QEvent, QRect, QTimer, QPropertyAnimation, QEasingCurve, pyqtSignal
)
from PyQt6.QtGui import QCursor, QColor

# key, caption, and the colours each block is drawn with, light and dark.
BLOCK_SPECS = [
    ("Carbonate", "Carbonate", "#DAE9F8", "#8FB4D4", "#2C4A63", "#5B8DB8"),
    ("Yield",     "Yield",     "#FFE0BF", "#DBA86A", "#5A4327", "#B98A4E"),
    ("CO2",       "CO₂",  "#E2EFDA", "#9CC08A", "#33503B", "#6A9A5B"),
]
DATA_BLOCKS = ("Carbonate", "CO2")

DRAG_THRESHOLD = 5      # pixels before a press becomes a drag
FONT_PX = 12            # caption size: small enough that "Carbonate" always fits
GRIP_WIDTH = 10
MARK_WIDTH = 14
SEGMENT_HEIGHT = 34
TRACK_PAD = 4           # inside the rounded track
SEGMENT_GAP = 4
SLIDE_MS = 170          # how long a segment takes to slide to its slot


class BlockSegment(QFrame):
    """One block inside the control."""

    def __init__(self, key, caption, spec, bar):
        super().__init__(bar.track)
        self.key = key
        self.caption = caption
        self.spec = spec
        self.bar = bar
        self.on = True
        self.hovered = False
        self.dragging = False
        self._press_pos = None
        self._grab_dx = 0

        self.setCursor(QCursor(Qt.CursorShape.OpenHandCursor))
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

        row = QHBoxLayout(self)
        row.setContentsMargins(6, 0, 6, 0)
        row.setSpacing(0)

        # The grip is the "you can move this" cue; without it a segment reads
        # as a button you can only click.
        self.grip = QLabel("⠿")
        self.grip.setFixedWidth(GRIP_WIDTH)
        row.addWidget(self.grip)
        row.addStretch()
        # The mark gets a label of its own: sharing one with the caption meant
        # a long caption pushed it past the edge and clipped it away.
        self.mark = QLabel()
        self.mark.setFixedWidth(MARK_WIDTH)
        self.mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        row.addWidget(self.mark)
        row.addSpacing(4)
        self.label = QLabel(caption)
        self.label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        row.addWidget(self.label)
        row.addStretch()
        row.addSpacing(GRIP_WIDTH)   # balances the grip so the caption stays centred

        self.slide = QPropertyAnimation(self, b"geometry", self)
        self.slide.setDuration(SLIDE_MS)
        self.slide.setEasingCurve(QEasingCurve.Type.OutCubic)

    # -- appearance ----------------------------------------------------
    def set_on(self, on):
        self.on = bool(on)
        self.restyle()

    def restyle(self):
        light_fill, light_edge, dark_fill, dark_edge = self.spec
        dark = self.bar.is_dark()
        if self.on:
            fill = dark_fill if dark else light_fill
            edge = dark_edge if dark else light_edge
            text = "#E8EAED" if dark else "#1F2933"
            grip = "#9FB3C8" if dark else "#7E93A8"
            mark, weight = "✓", "bold"
        else:
            fill = "#26282B" if dark else "#ECEDEF"
            edge = "#3A3D42" if dark else "#DCDEE1"
            text = "#7C848F" if dark else "#9AA1AB"
            grip = "#4A4F56" if dark else "#C2C7CE"
            mark, weight = "+", "normal"

        width = 1
        if self.dragging:
            edge, width = "#4C6EF5", 2
        elif self.hovered:
            edge = dark_edge if dark else light_edge

        self.setStyleSheet(f"""
            BlockSegment {{
                background-color: {fill};
                border: {width}px solid {edge};
                border-radius: 5px;
            }}
        """)
        self.label.setStyleSheet(
            f"color: {text}; font-size: {FONT_PX}px; font-weight: {weight}; background: transparent;")
        self.mark.setStyleSheet(
            f"color: {text}; font-size: {FONT_PX}px; font-weight: bold; background: transparent;")
        self.grip.setStyleSheet(f"color: {grip}; font-size: 12px; background: transparent;")
        self.mark.setText(mark)

    def lift(self, up):
        """Picks the segment up off the track, or sets it back down."""
        self.dragging = up
        if up:
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(18)
            shadow.setOffset(0, 3)
            shadow.setColor(QColor(0, 0, 0, 110))
            self.setGraphicsEffect(shadow)
            self.raise_()
            self.setCursor(QCursor(Qt.CursorShape.ClosedHandCursor))
        else:
            self.setGraphicsEffect(None)
            self.setCursor(QCursor(Qt.CursorShape.OpenHandCursor))
        self.restyle()

    def glide_to(self, rect, animate=True):
        self.slide.stop()
        if not animate or not self.isVisible():
            self.setGeometry(rect)
            return
        if self.geometry() == rect:
            return
        self.slide.setStartValue(self.geometry())
        self.slide.setEndValue(rect)
        self.slide.start()

    # -- hover ---------------------------------------------------------
    def enterEvent(self, event):
        self.hovered = True
        self.restyle()

    def leaveEvent(self, event):
        self.hovered = False
        self.restyle()

    # -- click vs drag -------------------------------------------------
    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        self._press_pos = event.globalPosition().toPoint()
        self._grab_dx = int(event.position().x())
        self.dragging = False

    def mouseMoveEvent(self, event):
        if self._press_pos is None:
            return
        pos = event.globalPosition().toPoint()
        if not self.dragging:
            if (pos - self._press_pos).manhattanLength() < DRAG_THRESHOLD:
                return
            self.slide.stop()
            self.lift(True)
        self.bar.drag_moved(self, pos, self._grab_dx)

    def mouseReleaseEvent(self, event):
        if self._press_pos is None:
            return
        was_drag = self.dragging
        self._press_pos = None
        if was_drag:
            self.lift(False)
            self.bar.drag_finished()
        else:
            self.bar.toggle(self)


class BlockLayoutBar(QWidget):
    """The segmented control, and the rules that keep the arrangement valid."""

    changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(5)

        self.track = QFrame()
        self.track.setObjectName("blockTrack")
        self.track.setFixedHeight(SEGMENT_HEIGHT + TRACK_PAD * 2)
        # Just enough for the longest caption in each of the three segments;
        # past that the control takes whatever width it is given.
        self.track.setMinimumWidth(340)
        self.track.installEventFilter(self)
        outer.addWidget(self.track)

        self.hint = QLabel()
        self.hint.setWordWrap(True)
        self.hint.setMinimumWidth(1)
        outer.addWidget(self.hint)

        # Source of truth for the arrangement: no layout manages the segments,
        # so a dragged one can sit wherever the mouse has it.
        self._order = []
        for key, caption, lf, le, df, de in BLOCK_SPECS:
            self._order.append(BlockSegment(key, caption, (lf, le, df, de), self))
        self.segments = list(self._order)

        self.set_state(["Carbonate", "Yield", "CO2"],
                       {"Carbonate": True, "Yield": False, "CO2": False})

    # -- geometry ------------------------------------------------------
    def eventFilter(self, obj, event):
        if obj is self.track and event.type() == QEvent.Type.Resize:
            self._place(animate=False)
        return super().eventFilter(obj, event)

    def _slot_rect(self, index):
        total = self.track.width() - TRACK_PAD * 2
        count = len(self._order)
        span = (total - SEGMENT_GAP * (count - 1)) / count
        x = TRACK_PAD + index * (span + SEGMENT_GAP)
        return QRect(round(x), TRACK_PAD, round(span), SEGMENT_HEIGHT)

    def _place(self, animate=True, skip=None):
        for i, seg in enumerate(self._order):
            if seg is skip:
                continue
            seg.glide_to(self._slot_rect(i), animate=animate)

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh_theme()
        self._place(animate=False)

    # -- theme ---------------------------------------------------------
    def is_dark(self):
        """Follows the window's own dark-mode flag, falling back to its palette."""
        window = self.window()
        if hasattr(window, "dark_mode"):
            return bool(window.dark_mode)
        return self.palette().window().color().lightness() < 128

    def refresh_theme(self):
        dark = self.is_dark()
        self.track.setStyleSheet(f"""
            QFrame#blockTrack {{
                background-color: {"#17181A" if dark else "#FFFFFF"};
                border: 1px solid {"#3A3B3E" if dark else "#DADCE0"};
                border-radius: 8px;
            }}
        """)
        self._restore_hint()
        for seg in self.segments:
            seg.restyle()

    # -- state ---------------------------------------------------------
    def order(self):
        """Block keys, left to right."""
        return [s.key for s in self._order]

    def enabled(self):
        return {s.key: s.on for s in self.segments}

    def set_state(self, order, enabled):
        order = [k for k in (order or []) if any(s.key == k for s in self.segments)]
        for key, *_ in BLOCK_SPECS:
            if key not in order:
                order.append(key)
        by_key = {s.key: s for s in self.segments}
        self._order = [by_key[k] for k in order]
        for seg in self._order:
            seg.set_on(enabled.get(seg.key, False))
        if not any(by_key[k].on for k in DATA_BLOCKS):
            by_key["Carbonate"].set_on(True)
        self._enforce_rules()
        self._update_hint()
        self._place(animate=False)

    # -- interaction ---------------------------------------------------
    def toggle(self, seg):
        if seg.on and seg.key in DATA_BLOCKS:
            others = [s for s in self.segments if s.key in DATA_BLOCKS and s is not seg]
            if not any(s.on for s in others):
                # The last data block cannot be dropped - there would be
                # nothing left to normalize.
                self._say("Carbonate or CO₂ has to stay in — one of them must run.")
                return
        seg.set_on(not seg.on)
        self._enforce_rules()
        self._update_hint()
        self._place()
        self.changed.emit()

    def drag_moved(self, seg, global_pos, grab_dx):
        """Keeps the segment under the mouse and opens a gap where it will land."""
        rect = self._slot_rect(0)
        x = self.track.mapFromGlobal(global_pos).x() - grab_dx
        x = max(TRACK_PAD, min(x, self.track.width() - TRACK_PAD - rect.width()))
        seg.setGeometry(QRect(int(x), TRACK_PAD, rect.width(), SEGMENT_HEIGHT))

        # Which slot is it hovering over? Measured from its centre, so the
        # swap happens as soon as it has genuinely passed a neighbour.
        span = rect.width() + SEGMENT_GAP
        target = round((seg.geometry().center().x() - TRACK_PAD - rect.width() / 2) / span)
        target = max(0, min(int(target), len(self._order) - 1))
        if seg.key not in DATA_BLOCKS:
            target = max(1, target)      # Yield can never take the first slot
        if target != self._order.index(seg):
            self._order.remove(seg)
            self._order.insert(target, seg)
            self._update_hint()
            self._place(skip=seg)        # everything else slides aside

    def drag_finished(self):
        self._enforce_rules()
        self._update_hint()
        self._place()
        self.changed.emit()

    # -- rules ---------------------------------------------------------
    def _enforce_rules(self):
        """Keeps Yield out of the first slot, and off the front of the run.

        The second half matters because a dropped block still holds its place:
        with Carbonate out, Yield sitting second would otherwise lead the
        blocks that actually get drawn.
        """
        for _ in range(len(self.segments)):
            if not self._order:
                return
            live = [s for s in self._order if s.on]
            offender = anchor = None
            if self._order[0].key not in DATA_BLOCKS:
                offender = self._order[0]
                anchor = next((s for s in self._order if s.key in DATA_BLOCKS), None)
            elif live and live[0].key not in DATA_BLOCKS:
                offender = live[0]
                # Has to land behind a block that is actually drawn, not just
                # behind any data block: a dropped one would leave it leading.
                anchor = next((s for s in live if s.key in DATA_BLOCKS), None)
            if offender is None or anchor is None:
                return
            self._order.remove(offender)
            self._order.insert(self._order.index(anchor) + 1, offender)

    # -- hint ----------------------------------------------------------
    def _update_hint(self):
        live = [s.caption for s in self._order if s.on]
        self.hint.setText(
            f"Sheet order: {'  ›  '.join(live)}"
            "   ·   click to include or drop, drag to reorder")

    def _say(self, message):
        dark = self.is_dark()
        self.hint.setText(message)
        self.hint.setStyleSheet(
            f"color: {'#E57373' if dark else '#C0392B'}; font-size: 12px; font-weight: bold;")
        QTimer.singleShot(2400, self._restore_hint)

    def _restore_hint(self):
        dark = self.is_dark()
        self.hint.setStyleSheet(
            f"color: {'#8B929C' if dark else '#888'}; font-style: italic; font-size: 12px;")
        self._update_hint()
