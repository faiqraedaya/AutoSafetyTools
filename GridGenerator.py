import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QSpinBox, QComboBox, QPushButton, QColorDialog, QFileDialog)
from PyQt5.QtGui import QPainter, QPen, QColor, QImage, QPixmap
from PyQt5.QtCore import Qt

class GridGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Grid Image Generator')
        self.setGeometry(100, 100, 1200, 800)

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)

        # Left panel for controls
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        left_panel.setFixedWidth(300)

        # Dimensions
        dim_group = QWidget()
        dim_layout = QVBoxLayout()
        dim_group.setLayout(dim_layout)
        self.width_spin = QSpinBox()
        self.width_spin.setRange(100, 5000)
        self.width_spin.setValue(1000)
        self.height_spin = QSpinBox()
        self.height_spin.setRange(100, 5000)
        self.height_spin.setValue(1000)
        dim_layout.addWidget(QLabel('Width:'))
        dim_layout.addWidget(self.width_spin)
        dim_layout.addWidget(QLabel('Height:'))
        dim_layout.addWidget(self.height_spin)

        # Grid settings
        grid_group = QWidget()
        grid_layout = QVBoxLayout()
        grid_group.setLayout(grid_layout)
        self.grid_x_spin = QSpinBox()
        self.grid_x_spin.setRange(1, 100)
        self.grid_x_spin.setValue(10)
        self.grid_y_spin = QSpinBox()
        self.grid_y_spin.setRange(1, 100)
        self.grid_y_spin.setValue(10)
        grid_layout.addWidget(QLabel('Grid X:'))
        grid_layout.addWidget(self.grid_x_spin)
        grid_layout.addWidget(QLabel('Grid Y:'))
        grid_layout.addWidget(self.grid_y_spin)

        # Line settings
        line_group = QWidget()
        line_layout = QVBoxLayout()
        line_group.setLayout(line_layout)
        self.line_type = QComboBox()
        self.line_type.addItems(['Solid', 'Dashed', 'Dotted'])
        self.line_weight = QSpinBox()
        self.line_weight.setRange(1, 50)
        self.line_weight.setValue(1)
        self.line_color = QPushButton('Line Color')
        self.line_color.clicked.connect(self.choose_line_color)
        self.line_color_value = QColor(0, 0, 0)
        line_layout.addWidget(QLabel('Line Type:'))
        line_layout.addWidget(self.line_type)
        line_layout.addWidget(QLabel('Line Weight:'))
        line_layout.addWidget(self.line_weight)
        line_layout.addWidget(self.line_color)

        # Background color
        self.bg_color = QPushButton('Background Color')
        self.bg_color.clicked.connect(self.choose_bg_color)
        self.bg_color_value = QColor(255, 255, 255)

        # Add all groups to left panel
        left_layout.addWidget(dim_group)
        left_layout.addWidget(grid_group)
        left_layout.addWidget(line_group)
        left_layout.addWidget(self.bg_color)
        
        # Generate and Save buttons
        generate_btn = QPushButton('Generate Grid')
        generate_btn.clicked.connect(self.generate_grid)
        save_btn = QPushButton('Save Image')
        save_btn.clicked.connect(self.save_image)
        left_layout.addWidget(generate_btn)
        left_layout.addWidget(save_btn)
        left_layout.addStretch()

        # Right panel for preview
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        right_layout.addWidget(self.preview_label)

        # Add panels to main layout
        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel, stretch=1)

        self.image = None

    def choose_line_color(self):
        color = QColorDialog.getColor(self.line_color_value)
        if color.isValid():
            self.line_color_value = color
            self.line_color.setStyleSheet(f'background-color: {color.name()}')

    def choose_bg_color(self):
        color = QColorDialog.getColor(self.bg_color_value)
        if color.isValid():
            self.bg_color_value = color
            self.bg_color.setStyleSheet(f'background-color: {color.name()}')

    def generate_grid(self):
        width = self.width_spin.value()
        height = self.height_spin.value()
        
        self.image = QImage(width, height, QImage.Format_RGB32)
        self.image.fill(self.bg_color_value)
        
        painter = QPainter(self.image)
        pen = QPen(self.line_color_value)
        pen.setWidth(self.line_weight.value())
        
        if self.line_type.currentText() == 'Dashed':
            pen.setStyle(Qt.DashLine)
        elif self.line_type.currentText() == 'Dotted':
            pen.setStyle(Qt.DotLine)
            
        painter.setPen(pen)
        
        # Draw vertical lines
        cell_width = width / self.grid_x_spin.value()
        for i in range(self.grid_x_spin.value() + 1):
            x = i * cell_width
            painter.drawLine(int(x), 0, int(x), height)
            
        # Draw horizontal lines
        cell_height = height / self.grid_y_spin.value()
        for i in range(self.grid_y_spin.value() + 1):
            y = i * cell_height
            painter.drawLine(0, int(y), width, int(y))
            
        painter.end()

        # Update preview
        pixmap = QPixmap.fromImage(self.image)
        scaled_pixmap = pixmap.scaled(800, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.preview_label.setPixmap(scaled_pixmap)

    def save_image(self):
        if self.image is None:
            return
            
        filename, _ = QFileDialog.getSaveFileName(self, "Save Image",
                                                "", "PNG Files (*.png);;All Files (*)")
        if filename:
            self.image.save(filename)

def main():
    app = QApplication(sys.argv)
    ex = GridGenerator()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()