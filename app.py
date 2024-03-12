import sys
import os
import csv
import pygame
import smtplib
from email import encoders
from PyQt5.QtGui import QIcon
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from PyQt5.QtWidgets import QListWidget, QListWidgetItem
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QComboBox, QColorDialog, QFontDialog, QMessageBox, QInputDialog
)
from PyQt5.QtGui import QPixmap, QImage, QFont, QColor, QFontDatabase, QPainter, QFontMetrics
from PyQt5.QtCore import Qt

class CertificateGeneratorApp(QWidget):
    def generate_certificate_for_name(self, name):
        if not all([self.template_path, self.output_path]):
            QMessageBox.warning(self, "Warning", "Please fill in all required fields.")
            return

        # Pygame initialization
        if not self.pygame_initialized:
            pygame.init()
            self.pygame_initialized = True

        # Load the template image
        template_image = pygame.image.load(self.template_path)

        # Create a new surface to combine the template and certificate
        combined_surface = pygame.Surface(template_image.get_size(), pygame.SRCALPHA)

        # Blit the template onto the combined surface
        combined_surface.blit(template_image, (0, 0))

        # Generate Certificate Image
        certificate_surface = self.create_certificate(name)

        # Get the size of the certificate surface
        cert_width, cert_height = certificate_surface.get_size()

        # Calculate the position to center the certificate on the template
        cert_x = (combined_surface.get_width() - cert_width) // 2
        cert_y = (combined_surface.get_height() - cert_height) // 2

        # Blit the certificate (with the name) onto the combined surface
        combined_surface.blit(certificate_surface, (cert_x, cert_y))

        # Convert the combined Pygame Surface to a QPixmap for preview
        combined_pixmap = self.qimage_to_pixmap(combined_surface)

        # Display the combined certificate in the preview label
        self.preview_label.setPixmap(combined_pixmap)

        # Convert the combined surface to PNG (you can save this or convert to PDF)
        image_path = os.path.join(self.output_path, f"{name}_certificate.png")
        pygame.image.save(combined_surface, image_path)

        # Show success message
        QMessageBox.information(self, "Success", f"Certificate for {name} generated successfully.")

    def __init__(self):
        super().__init__()

        # Variables
        self.template_path = ""
        self.csv_path = ""
        self.output_path = ""
        self.name_column = ""
        self.email_column = ""
        self.x_position = 0
        self.y_position = 0
        self.selected_font = QFont()  # Initialize with default font
        self.text_color = (0, 0, 0)
        self.font_styles = []
        self.custom_subject = ""
        self.custom_message = ""
        self.email_status_list = QListWidget()
        self.status_window = None  # Initialize as an instance variable

        # Pygame variables
        self.pygame_initialized = False

        self.layout = QVBoxLayout()
        self.main_layout = QHBoxLayout()
        self.init_ui()

    def init_ui(self):
        # Logo
        logo_label = QLabel()
        logo_pixmap = QPixmap("\logo.png")  # Replace with the actual path to your logo
        logo_label.setPixmap(logo_pixmap)
        logo_label.setAlignment(Qt.AlignCenter)

        # Main Layout Configuration
        self.main_layout = QHBoxLayout()
        self.main_layout.addWidget(logo_label)
        self.main_layout.addStretch(1)  # Add a spacer item to center the logo
        self.main_layout.addLayout(self.layout)

        # Template Selection
        template_label = QLabel("Template PNG Path:")
        self.template_entry = QLineEdit()
        template_button = QPushButton("Browse")
        template_button.clicked.connect(self.browse_template)

        # CSV Selection
        csv_label = QLabel("CSV Path:")
        self.csv_entry = QLineEdit()
        csv_button = QPushButton("Browse")
        csv_button.clicked.connect(self.browse_csv)

        # Font Family Dropdown
        self.font_family_dropdown = QComboBox()
        self.load_system_fonts()
        self.font_family_dropdown.currentTextChanged.connect(self.update_font_family)

        # Font Style Checkboxes
        self.bold_checkbox = QPushButton("Bold")
        self.bold_checkbox.setCheckable(True)
        self.bold_checkbox.clicked.connect(self.update_preview)

        self.italic_checkbox = QPushButton("Italic")
        self.italic_checkbox.setCheckable(True)
        self.italic_checkbox.clicked.connect(self.update_preview)

        self.underline_checkbox = QPushButton("Underline")
        self.underline_checkbox.setCheckable(True)
        self.underline_checkbox.clicked.connect(self.update_preview)

        self.strikethrough_checkbox = QPushButton("Strikethrough")
        self.strikethrough_checkbox.setCheckable(True)
        self.strikethrough_checkbox.clicked.connect(self.update_preview)

        # Font Size Dropdown
        font_size_label = QLabel("Font Size:")
        self.font_size_dropdown = QComboBox()
        self.font_size_dropdown.addItems([str(i) for i in range(100000)])  # Add font sizes from 0 to 99999
        self.font_size_dropdown.setCurrentText(str(self.get_default_font_size()))  # Set default font size
        self.font_size_dropdown.currentTextChanged.connect(self.update_font_size)

        # Font Color Button
        text_color_label = QLabel("Text Color:")
        self.text_color_button = QPushButton("Choose Color")
        self.text_color_button.clicked.connect(self.choose_text_color)
        self.text_color_button.setStyleSheet(f"background-color: {self.rgb_to_hex(self.text_color)};")

        # X and Y Position
        x_position_label = QLabel("X Position:")
        self.x_position_entry = QLineEdit(str(self.x_position))
        self.x_position_entry.textChanged.connect(self.update_x_position)

        y_position_label = QLabel("Y Position:")
        self.y_position_entry = QLineEdit(str(self.y_position))
        self.y_position_entry.textChanged.connect(self.update_y_position)

        # CSV Columns Dropdowns
        name_column_label = QLabel("Name Column:")
        self.name_column_dropdown = QComboBox()
        self.name_column_dropdown.currentTextChanged.connect(self.update_name_column)

        email_column_label = QLabel("Email Column:")
        self.email_column_dropdown = QComboBox()
        self.email_column_dropdown.currentTextChanged.connect(self.update_email_column)

        # Image Preview
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)

        self.layout.addWidget(self.email_status_list)

        # Customize Email Button
        customize_button = QPushButton("Customize Email")
        customize_button.clicked.connect(self.customize_email)

        # Generate and Send Certificates Button
        generate_button = QPushButton("Generate Certificates")
        generate_button.clicked.connect(self.generate_certificates)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(template_label)
        layout.addWidget(self.template_entry)
        layout.addWidget(template_button)

        layout.addWidget(csv_label)
        layout.addWidget(self.csv_entry)
        layout.addWidget(csv_button)

        layout.addWidget(QLabel("Font Family:"))
        layout.addWidget(self.font_family_dropdown)

        layout.addWidget(QLabel("Font Styles:"))
        layout.addWidget(self.bold_checkbox)
        layout.addWidget(self.italic_checkbox)
        layout.addWidget(self.underline_checkbox)
        layout.addWidget(self.strikethrough_checkbox)

        layout.addWidget(font_size_label)
        layout.addWidget(self.font_size_dropdown)

        layout.addWidget(text_color_label)
        layout.addWidget(self.text_color_button)

        layout.addWidget(x_position_label)
        layout.addWidget(self.x_position_entry)

        layout.addWidget(y_position_label)
        layout.addWidget(self.y_position_entry)

        layout.addWidget(name_column_label)
        layout.addWidget(self.name_column_dropdown)

        layout.addWidget(email_column_label)
        layout.addWidget(self.email_column_dropdown)

        layout.addWidget(customize_button)
        layout.addWidget(generate_button)

        preview_layout = QVBoxLayout()
        preview_layout.addWidget(self.preview_label)

        main_layout = QHBoxLayout()
        main_layout.addLayout(layout)
        main_layout.addLayout(preview_layout)

        self.setLayout(main_layout)
        self.setWindowTitle("Certificate Generator")
        self.setGeometry(100, 100, 960, 720)  # Set the window size to 960x720

    def browse_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Template PNG", "", "PNG files (*.png);;All Files (*)")
        if file_path:
            self.template_entry.setText(file_path)
            self.template_path = file_path
            self.update_preview()

    def browse_csv(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV", "", "CSV files (*.csv);;All Files (*)")
        if file_path:
            self.csv_entry.setText(file_path)
            self.csv_path = file_path
            self.load_csv_columns()

    def choose_text_color(self):
        color = QColorDialog.getColor(initial=QColor(*self.text_color))
        if color.isValid():
            self.text_color = color.getRgb()[:3]
            self.text_color_button.setStyleSheet(f"background-color: {self.rgb_to_hex(self.text_color)};")
            self.update_preview()

    def update_x_position(self):
        try:
            self.x_position = int(self.x_position_entry.text())
            self.update_preview()
        except ValueError:
            pass

    def update_y_position(self):
        try:
            self.y_position = int(self.y_position_entry.text())
            self.update_preview()
        except ValueError:
            pass

    def update_font_family(self):
        font_family = self.font_family_dropdown.currentText()
        self.selected_font.setFamily(font_family)
        self.update_preview()

    def update_name_column(self, selected_text):
        self.name_column = selected_text
        self.update_preview()

    def update_email_column(self, selected_text):
        self.email_column = selected_text
        self.update_preview()

    def update_font_size(self):
        font_size = int(self.font_size_dropdown.currentText())
        self.selected_font.setPointSize(font_size-100)
        self.update_preview()
        self.selected_font.setPointSize(font_size)

    def update_preview(self):
        if self.template_path:
            image = QImage(self.template_path)
            pixmap = QPixmap.fromImage(image)

            painter = QPainter(pixmap)
            painter.setFont(self.selected_font)
            painter.setPen(QColor(*self.text_color))

            # Apply font styles
            font = self.selected_font
            font.setBold(self.bold_checkbox.isChecked())
            font.setItalic(self.italic_checkbox.isChecked())
            font.setUnderline(self.underline_checkbox.isChecked())
            font.setStrikeOut(self.strikethrough_checkbox.isChecked())
            painter.setFont(font)

            font_size = int(self.font_size_dropdown.currentText())
            # Display a sample name for preview
            sample_name = "Sample text"  # Replace with a name from your dataset
            # Adjust position calculation based on font metrics
            x_position = (pixmap.width() // 2)-(QFontMetrics(font).width(sample_name)//2) + self.x_position
            y_position = (pixmap.height()//2) - self.y_position+font_size

            painter.drawText(x_position, y_position, sample_name)
            painter.end()

            self.preview_label.setPixmap(pixmap.scaled(540, 540, Qt.KeepAspectRatio))

    def load_csv_columns(self):
        if self.csv_path:
            with open(self.csv_path, 'r') as csv_file:
                header = csv.reader(csv_file).__next__()
                header = [column.strip() for column in header]  # Strip whitespaces from column names
                self.name_column_dropdown.clear()
                self.email_column_dropdown.clear()
                self.name_column_dropdown.addItems(header)
                self.email_column_dropdown.addItems(header)

    def customize_email(self):
        custom_subject, ok_subject = QInputDialog.getText(self, "Custom Subject", "Enter custom subject:")
        custom_message, ok_message = QInputDialog.getMultiLineText(self, "Custom Message", "Enter custom message:")

        if ok_subject and ok_message:
            QMessageBox.information(self, "Success", "Customization successful.")
        else:
            QMessageBox.warning(self, "Warning", "Operation canceled by user.")

        self.custom_subject = custom_subject if ok_subject else ""
        self.custom_message = custom_message if ok_message else ""

    def generate_certificates(self):
        if not all([self.template_path, self.csv_path]):
            QMessageBox.warning(self, "Warning", "Please fill in all required fields.")
            return

        # Pygame initialization
        if not self.pygame_initialized:
            pygame.init()
            self.pygame_initialized = True

        # Loop through CSV and generate certificates
        with open(self.csv_path, 'r') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            # Ensure that the required columns exist in the CSV file
            if self.name_column not in csv_reader.fieldnames or self.email_column not in csv_reader.fieldnames:
                QMessageBox.warning(self, "Error", f"Name or Email column not found in CSV. Selected: {self.name_column}, {self.email_column}")
                return

            # Clear previous items in the list
            self.email_status_list.clear()

            for row in csv_reader:
                name = row[self.name_column]
                email = row[self.email_column]

                # Generate Certificate Image
                certificate_surface = self.create_certificate(name)

                # Convert Surface to PNG (you can save this or convert to PDF)
                image_path = os.path.join(self.output_path, f"{name}_certificate.png")
                pygame.image.save(certificate_surface, image_path)

                # Send Email with Certificate Image
                email_sent = self.send_email(email, image_path)

                # Delete temporary image file
                os.remove(image_path)

                # Update email status list
                item = QListWidgetItem(f"Email sent to {email}" if email_sent else f"Failed to send email to {email}")
                self.email_status_list.addItem(item)

            # Show the email status list in a new window
            self.status_window = QWidget()
            status_layout = QVBoxLayout()
            status_layout.addWidget(self.email_status_list)
            self.status_window.setLayout(status_layout)
            self.status_window.setWindowTitle("Email Status")
            self.status_window.show()

        QMessageBox.information(self, "Success", "Certificates generated successfully.")

    def create_certificate(self, name):
        # Check if the template path is set
        if not self.template_path:
            QMessageBox.warning(self, "Warning", "Please select a template.")
            return None

        # Load the template image using Pygame
        template_image = pygame.image.load(self.template_path)

        # Create a copy of the template image to avoid modifying the original
        certificate_surface = template_image.copy()

        # Convert the certificate surface to a Pygame surface
        certificate_surface = pygame.surfarray.array3d(certificate_surface)
        certificate_surface = pygame.surfarray.make_surface(certificate_surface)

        font = pygame.font.SysFont(self.selected_font.family(), self.selected_font.pointSize())
        text_color = self.text_color

        # Apply font styles
        font.set_bold(self.bold_checkbox.isChecked())
        font.set_italic(self.italic_checkbox.isChecked())
        font.set_underline(self.underline_checkbox.isChecked())

        # Calculate the text position based on the center origin
        text_x = (certificate_surface.get_width() // 2)-(font.size(name)[0]//2)+ self.x_position
        text_y = (certificate_surface.get_height() // 2)- self.y_position 

        # Render and blit the text on the template surface
        text_surface = font.render(name, True, text_color)
        certificate_surface.blit(text_surface, (text_x, text_y))

        # Check for strikethrough and draw a line through the text
        if self.strikethrough_checkbox.isChecked():
            line_height = font.get_linesize() // 2
            pygame.draw.line(certificate_surface, text_color, (text_x, text_y + line_height),
                             (text_x + text_surface.get_width(), text_y + line_height), 2)

        return certificate_surface

    def get_default_font_size(self):
        if self.template_path:
            image = QImage(self.template_path)
            template_height = image.height()
            default_font_size = max(1, template_height // 7)  # Ensure font size is at least 1
            return default_font_size
        else:
            return 200  # Default font size if template height is not available

    def load_system_fonts(self):
        # Fetch the list of available font families
        font_database = QFontDatabase()
        font_families = font_database.families(QFontDatabase.Any)

        # Populate the font family dropdown
        self.font_family_dropdown.addItems(font_families)

    def qimage_to_pixmap(self, image_surface):
        # Convert Pygame Surface to QPixmap
        width, height = image_surface.get_size()
        image_data = pygame.image.tostring(image_surface, 'RGBA')
        image_qt = QImage(image_data, width, height, QImage.Format_RGBA8888)

        pixmap = QPixmap.fromImage(image_qt)
        return pixmap

    def send_email(self, recipient_email, image_path):
            # Gmail SMTP setup
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            smtp_username = 'example@gmail.com'
            smtp_password = 'password'#app Password

            # Email configuration
            sender_email = 'omborekar18@gmail.com'
            subject = f'Certificate: {self.custom_subject}'
            body = f'{self.custom_message}\n\n©Certificate Generator & Automation'

            # Create MIME object
            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = recipient_email
            message['Subject'] = subject
            message.attach(MIMEText(body, 'plain'))

            # Attach PNG file
            with open(image_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(image_path)}')
                encoders.encode_base64(part)
                message.attach(part)

            try:
                # Connect to Gmail SMTP server and send email
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(smtp_username, smtp_password)
                    server.sendmail(sender_email, recipient_email, message.as_string())

                return True  # Email sent successfully
            except smtplib.SMTPException as e:
                print(f"SMTP Exception: {e}")
                return False  # Failed to send email
            except Exception as e:
                print(f"Error sending email to {recipient_email}: {e}")
                print(traceback.format_exc())  # Print detailed error information
                return False  # Failed to send email

    def rgb_to_hex(self, rgb):
        return "#{:02x}{:02x}{:02x}".format(rgb[0], rgb[1], rgb[2])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CertificateGeneratorApp()
    window.show()
    sys.exit(app.exec_())
