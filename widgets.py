from PyQt5.QtWidgets import QComboBox, QStyledItemDelegate, QStyle
from PyQt5.QtCore import Qt, QEvent, QVariantAnimation, QRect, QPointF, QSize
from PyQt5.QtGui import QPainter, QColor, QFont

class ModernComboBox(QComboBox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setFixedHeight(44)
        self.setCursor(Qt.PointingHandCursor)
        self.view().setStyleSheet("background: transparent; border: none;")
        self.view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint)
        self.view().window().setAttribute(Qt.WA_TranslucentBackground)
        self.arrow_angle = 0
        self.animation = QVariantAnimation(startValue=0, endValue=180, duration=220, valueChanged=self.on_animation)
        self.installEventFilter(self)

    def on_animation(self, value):
        self.arrow_angle = value
        self.update()

    def showPopup(self):
        self.animation.setDirection(QVariantAnimation.Forward)
        self.animation.start()
        super().showPopup()

    def hidePopup(self):
        self.animation.setDirection(QVariantAnimation.Backward)
        self.animation.start()
        super().hidePopup()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = QRect(0, 0, self.width(), self.height())
        painter.setBrush(QColor("#FFFFFF"))
        pen_color = QColor("#1976D2") if self.hasFocus() else QColor("#C5CAE9")
        painter.setPen(pen_color)
        painter.drawRoundedRect(rect.adjusted(1,1,-1,-1),12,12)
        # Texte et flèche
        painter.setPen(QColor("#2E3B55"))
        text_rect = rect.adjusted(16,0,-44,0)
        current = self.currentText() or "Choisir une catégorie"
        painter.drawText(text_rect, Qt.AlignVCenter, current)
        arrow_x, arrow_y = self.width()-30, self.height()//2
        painter.save()
        painter.translate(arrow_x, arrow_y)
        painter.rotate(self.arrow_angle)
        painter.translate(-arrow_x, -arrow_y)
        painter.setBrush(QColor("#1976D2"))
        painter.setPen(Qt.NoPen)
        painter.drawPolygon(QPointF(arrow_x-6, arrow_y-4), QPointF(arrow_x+6, arrow_y-4), QPointF(arrow_x, arrow_y+4))
        painter.restore()

    def eventFilter(self, obj, event):
        if event.type() in (QEvent.Enter, QEvent.Leave):
            self.setCursor(Qt.PointingHandCursor if event.type()==QEvent.Enter else Qt.ArrowCursor)
        return super().eventFilter(obj, event)

class StyledItemDelegate(QStyledItemDelegate):
    icons = {
        "MATERIELS INFORMATIQUES": "💻",
        "FOURNITURES DE BUREAUX": "📋",
        "PRODUITS D'ENTRETIEN MENNAGER": "🧹",
        "HABILLEMENTS": "👔",
        "MOBILIER DE BUREAU": "🪑",
        "PARC AUTO": "🚗",
        "CONFECTION DES FOURNITURS IMPRIMEES": "🖨️",
        "CONSOMMABLE INFORMATIQUE": "🖱️",
        "PRODUITS PHARMACEUTIQUES": "💊",
        "EAUX": "💧",
        "": "📌"
    }

    def paint(self, painter, option, index):
        painter.setRenderHint(QPainter.Antialiasing)
        rect = option.rect.adjusted(4,4,-4,-4)
        painter.setBrush(QColor("#1976D2") if option.state & QStyle.State_Selected else QColor("#FFFFFF"))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(rect,12,12)
        text = index.data()
        icon = self.icons.get(text,"📦")
        painter.setFont(QFont("Segoe UI Emoji",18))
        painter.setPen(Qt.white if option.state & QStyle.State_Selected else QColor("#555"))
        painter.drawText(QRect(rect.left()+12, rect.top()+8,32,32), Qt.AlignCenter, icon)
        painter.setFont(QFont("Segoe UI",11,QFont.DemiBold))
        painter.setPen(Qt.white if option.state & QStyle.State_Selected else QColor("#2E3B55"))
        painter.drawText(rect.adjusted(56,8,-12,-8), Qt.AlignVCenter, text)

    def sizeHint(self, option, index):
        return QSize(200,56)
