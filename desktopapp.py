import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QTableWidget, QVBoxLayout, 
    QWidget, QTableWidgetItem, QLineEdit, QGridLayout, QHBoxLayout, 
    QSizePolicy, QSpacerItem, QScrollArea, QGroupBox, QComboBox, QCheckBox, QShortcut, QMessageBox, QLabel,
    QFrame
)
from PyQt5.QtGui import QKeySequence, QPixmap, QColor
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTableWidgetItem, QCheckBox, QMessageBox, QDialog, QDialogButtonBox
from PyQt5 import QtGui
import csv
import pyodbc
import requests
import configparser
import os


class SelecionarOpcao(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QtGui.QIcon('csv/favicon.ico'))
        self.setWindowTitle("Opção de Visualização")
        self.layout = QVBoxLayout(self)

        self.checkbox_cont = QCheckBox("Conteúdos")
        self.checkbox_prec = QCheckBox("Alteração de preços")

        self.layout.addWidget(self.checkbox_cont)
        self.layout.addWidget(self.checkbox_prec)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.layout.addWidget(self.button_box)

        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self.setLayout(self.layout)

        self.checkbox_cont.clicked.connect(self.checkbox_cont_clicked)
        self.checkbox_prec.clicked.connect(self.checkbox_preco_clicked)

        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

    def checkbox_cont_clicked(self):
        if self.checkbox_cont.isChecked():
            self.checkbox_prec.setChecked(False)

    def checkbox_preco_clicked(self):
        if self.checkbox_prec.isChecked():
            self.checkbox_cont.setChecked(False)

    def obter_selecao(self):
        if self.checkbox_cont.isChecked():
            return "conteudos"
        elif self.checkbox_prec.isChecked():
            return "alteracao_precos"
        else:
            return None

class JanelaComparacao(QDialog):
    def __init__(self, column1_values, column2_values, descricao_aq_values, ref_aq_values, column1_name, column2_name, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QtGui.QIcon('csv/favicon.ico'))
        self.setWindowTitle("Resultado da Comparação")
        self.column1_values = column1_values
        self.column2_values = column2_values
        self.descricao_aq_values = descricao_aq_values
        self.ref_aq_values = ref_aq_values
        self.column1_name = column1_name
        self.column2_name = column2_name
        self.current_product_index = 0

        main_layout = QVBoxLayout(self)
        self.setLayout(main_layout)

        app_bar = QHBoxLayout()
        self.download_button = QPushButton("Download imagem - F3")
        self.download_button.clicked.connect(self.download_imagem_atual)
        app_bar.addWidget(self.download_button)
        main_layout.addLayout(app_bar)

        key_instructions_label = QLabel("W - Anterior | S - Posterior")
        main_layout.addWidget(key_instructions_label)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumHeight(900)
        self.scroll_area.setMinimumWidth(1200)
        main_layout.addWidget(self.scroll_area)

        self.comparison_widget = QWidget()
        self.comparison_layout = QVBoxLayout(self.comparison_widget)

        self.scroll_area.setWidget(self.comparison_widget)

        self.current_image_url = None

        self.mostrar_produto_atual()

        flags = self.windowFlags() | Qt.WindowMinimizeButtonHint
        self.setWindowFlags(flags)

        self.default_path = self.obter_path()

    def obter_path(self):
        config = configparser.ConfigParser()
        try:
            config.read("csv/config.ini", encoding='utf-8')
            default_path = config.get("Settings", "default_path")
            return default_path
        except Exception as e:
            print(f"Erro a ler config.ini: {e}")
            return None

    def mostrar_produto_atual(self):
        # Limpar o layout atual
        for i in reversed(range(self.comparison_layout.count())):
            widget = self.comparison_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)

        if 0 <= self.current_product_index < len(self.column1_values):
            column1_value = self.column1_values[self.current_product_index]
            column2_value = self.column2_values[self.current_product_index]
            descricao_aq_value = self.descricao_aq_values[self.current_product_index]
            ref_aq_value = self.ref_aq_values[self.current_product_index]

            self.setWindowTitle(f"Resultado da Comparação: {ref_aq_value} - {descricao_aq_value}")

            header_label1 = QLabel(self.column1_name)
            header_label1.setStyleSheet("font-weight: bold; text-transform: uppercase;")
            self.comparison_layout.addWidget(header_label1)

            if 'Image' in self.column1_name and isinstance(column1_value, str) and column1_value.startswith('http'):
                image_label1 = QLabel()
                pixmap1 = self.carregar_imagem(column1_value)
                if pixmap1:
                    scaled_pixmap1 = pixmap1.scaledToHeight(380, Qt.SmoothTransformation)
                    image_label1.setPixmap(scaled_pixmap1)
                    self.current_image_url = column1_value
                self.comparison_layout.addWidget(image_label1)
            else:
                value_label1 = QLabel(str(column1_value))
                self.comparison_layout.addWidget(value_label1)

            frame1 = QFrame()
            frame1.setFrameShape(QFrame.HLine)
            frame1.setFrameShadow(QFrame.Sunken)
            frame1.setStyleSheet("QFrame { border: 2px solid red; }")
            self.comparison_layout.addWidget(frame1)

            header_label2 = QLabel(self.column2_name)
            header_label2.setStyleSheet("font-weight: bold; text-transform: uppercase;")
            self.comparison_layout.addWidget(header_label2)

            if 'Image' in self.column2_name and isinstance(column2_value, str) and column2_value.startswith('http'):
                image_label2 = QLabel()
                pixmap2 = self.carregar_imagem(column2_value)
                if pixmap2:
                    scaled_pixmap2 = pixmap2.scaledToHeight(380, Qt.SmoothTransformation)
                    image_label2.setPixmap(scaled_pixmap2)
                    self.current_image_url = column2_value
                self.comparison_layout.addWidget(image_label2)
            else:
                value_label2 = QLabel(str(column2_value))
                self.comparison_layout.addWidget(value_label2)

            frame2 = QFrame()
            frame2.setFrameShape(QFrame.HLine)
            frame2.setFrameShadow(QFrame.Sunken)
            frame2.setStyleSheet("QFrame { border: 2px solid red; }")
            self.comparison_layout.addWidget(frame2)

            self.scroll_area.ensureVisible(0, 0)
            self.scroll_area.verticalScrollBar().setValue(0)

    def carregar_imagem(self, url):
        try:
            response = requests.get(url)
            response.raise_for_status()
            image_data = response.content
            pixmap = QPixmap()
            pixmap.loadFromData(image_data)
            return pixmap
        except requests.RequestException as e:
            print(f"Carregamento da imagem falhou {url}: {e}")
            return None

    def download_imagem_atual(self):
        if not self.current_image_url or self.current_image_url != self.column2_values[self.current_product_index]:
            print("Não há imagem válida para fazer download na segunda coluna.")
            return

        try:
            response = requests.get(self.current_image_url)
            response.raise_for_status()
            image_data = response.content

            ref_aq_value = self.ref_aq_values[self.current_product_index]

            filename, _ = QFileDialog.getSaveFileName(self, "Guardar como", os.path.join(self.default_path, f"{ref_aq_value}.jpg"), "Imagem (*.jpg)")

            if not filename:
                print("Caminho não escolhido.")
                return

            with open(filename, 'wb') as f:
                f.write(image_data)

            print(f"Imagem guardada em: {filename}")
        except requests.RequestException as e:
            print(f"Erro ao fazer download: {e}")

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_S:
            self.current_product_index = (self.current_product_index + 1) % len(self.column1_values)
            self.mostrar_produto_atual()
        elif event.key() == Qt.Key_W:
            self.current_product_index = (self.current_product_index - 1) % len(self.column1_values)
            self.mostrar_produto_atual()
        elif event.key() == Qt.Key_F3:
            self.download_imagem_atual()
        else:
            super().keyPressEvent(event)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QtGui.QIcon('csv/favicon.ico'))
        self.setWindowTitle("Comparador")
        self.setGeometry(100, 100, 1000, 600)

        self.central_widget = QWidget() 
        self.setCentralWidget(self.central_widget)

        self.layout = QHBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Left Container (Table)
        self.left_container = QWidget()
        self.left_layout = QVBoxLayout()
        self.left_container.setLayout(self.left_layout)

        self.btn_load = QPushButton("Abrir ficheiro")
        self.btn_load.setStyleSheet("font-size: 16px;")
        self.btn_load.setMinimumSize(200, 50)
        self.btn_load.clicked.connect(self.carregar_ficheiro)
        self.left_layout.addWidget(self.btn_load)

        self.scroll_area = QScrollArea()
        self.left_layout.addWidget(self.scroll_area)

        self.table_widget = QTableWidget()
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.table_widget)

        self.layout.addWidget(self.left_container)

        # Set fixed size for the left container
        left_container_width = 1300
        self.left_container.setFixedWidth(left_container_width)

        # Create the combined container for right and bottom containers
        self.right_bottom_container = QGroupBox()  # Group box for the layout
        self.right_bottom_layout = QVBoxLayout()
        self.right_bottom_container.setLayout(self.right_bottom_layout)

        # Right Container (Filters and Buttons)
        self.right_container = QWidget()
        self.right_layout = QVBoxLayout()
        self.right_container.setLayout(self.right_layout)

        # Group Box for Filters
        self.filter_groupbox = QGroupBox("")
        self.right_layout.addWidget(self.filter_groupbox)
        self.filter_groupbox_layout = QVBoxLayout()
        self.filter_groupbox.setLayout(self.filter_groupbox_layout)

        self.filter_widgets = []
        self.data = None

        self.filter_layout = QGridLayout()
        self.filter_groupbox_layout.addLayout(self.filter_layout)

        # Add buttons to the group box
        self.button_layout = QHBoxLayout()

        # Add "Aplicar" button
        self.btn_apply_filter = QPushButton("Aplicar")
        self.btn_apply_filter.setStyleSheet(
            "background-color: #32612D; color: white; font-size: 16px;")
        self.btn_apply_filter.clicked.connect(self.apply_filter)
        self.btn_apply_filter.setVisible(False)
        self.button_layout.addWidget(self.btn_apply_filter)

        # Add "Limpar" button
        self.btn_clear_filters = QPushButton("Limpar")
        self.btn_clear_filters.setStyleSheet(
            "background-color : #710C04; color: white; font-size: 16px;")
        self.btn_clear_filters.clicked.connect(self.clear_filters)
        self.btn_clear_filters.setVisible(False)
        self.button_layout.addWidget(self.btn_clear_filters)

        # Add a vertical spacer to lower the position of the buttons
        spacer_item = QSpacerItem(20, 50, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.filter_groupbox_layout.addItem(spacer_item)

        # Add buttons layout to the group box layout
        self.filter_groupbox_layout.addLayout(self.button_layout)

        # Adjust the height of the group box 
        new_height = 450  # Set the desired height  
        self.filter_groupbox.setFixedHeight(new_height)

        # Adjust the maximum width of the group box
        max_width = 570  # Set the desired maximum width
        self.filter_groupbox.setMaximumWidth(max_width)
        
        # ----------------------------------------------------------------

        # Second Container (center_container)
        self.center_container = QWidget()
        self.center_layout = QVBoxLayout()
        self.center_container.setLayout(self.center_layout)

        # Group Box for Filters
        self.db2_groupbox = QGroupBox("")
        self.center_layout.addWidget(self.db2_groupbox)
        self.db2_groupbox_layout = QVBoxLayout()
        self.db2_groupbox.setLayout(self.db2_groupbox_layout)

        # Adjust the height of the group box 
        new_heightb = 200  # Set the desired height  
        self.db2_groupbox.setFixedHeight(new_heightb)

        # Adjust the maximum width of the group box
        self.db2_groupbox.setMaximumWidth(max_width)

        self.select_all_checkbox = QCheckBox("Selecionar todos os artigos")
        self.select_all_checkbox.setStyleSheet("font-size: 14px;")
        self.select_all_checkbox.setVisible(False)
        self.db2_groupbox_layout.addWidget(self.select_all_checkbox)

        # Connect the clicked signal of "Select All" checkbox
        self.select_all_checkbox.clicked.connect(self.selecionar_todos)

        # Add dropdown menus for column selection
        self.dropdown_layout = QVBoxLayout()

        # Dropdown menu for first column selection
        self.combo_box1 = QComboBox()
        self.combo_box1.setVisible(False)
        self.dropdown_layout.addWidget(self.combo_box1)

        # Dropdown menu for second column selection
        self.combo_box2 = QComboBox()
        self.combo_box2.setVisible(False)
        self.dropdown_layout.addWidget(self.combo_box2)

        # Add dropdown layout to the group box layout
        self.db2_groupbox_layout.addLayout(self.dropdown_layout)

        # Add "Comparar" button
        self.comparar_btn = QPushButton("Comparar")
        self.comparar_btn.setStyleSheet(
            "background-color: #6495ED; color: white; font-size: 16px;")
        self.comparar_btn.setVisible(False)
        self.comparar_btn.clicked.connect(self.comparar_produtos_selecionados)
        
        self.db2_groupbox_layout.addWidget(self.comparar_btn)

        # Add the center container to the layout
        self.right_bottom_layout.addWidget(self.center_container)

        # ----------------------------------------------------------------

        # Bottom Container (Third container in the right)
        self.bottom_container = QWidget()
        self.bottom_layout = QVBoxLayout()
        self.bottom_container.setLayout(self.bottom_layout)

        # Group Box for Filters
        self.db_groupbox = QGroupBox("")
        self.bottom_layout.addWidget(self.db_groupbox)
        self.db_groupbox_layout = QVBoxLayout()
        self.db_groupbox.setLayout(self.db_groupbox_layout)

        # Adjust the height of the group box 
        new_heightb = 200  # Set the desired height  
        self.db_groupbox.setFixedHeight(new_heightb)

        # Adjust the maximum width of the group box
        self.db_groupbox.setMaximumWidth(max_width)

        # Add checkboxes to the group box
        self.checkbox1 = QCheckBox("Opção 1")
        self.checkbox2 = QCheckBox("Opção 2")
        self.checkbox3 = QCheckBox("Opção 3")
        self.checkbox4 = QCheckBox("Opção 4")
        self.checkbox5 = QCheckBox("Opção 5")
        self.checkbox6 = QCheckBox("Opção 6")

        self.checkbox1.setVisible(False)
        self.checkbox2.setVisible(False)
        self.checkbox3.setVisible(False)
        self.checkbox4.setVisible(False)
        self.checkbox5.setVisible(False)
        self.checkbox6.setVisible(False)

        self.db_groupbox_layout.addWidget(self.checkbox1)
        self.db_groupbox_layout.addWidget(self.checkbox2)
        self.db_groupbox_layout.addWidget(self.checkbox3)
        self.db_groupbox_layout.addWidget(self.checkbox4)
        self.db_groupbox_layout.addWidget(self.checkbox5)
        self.db_groupbox_layout.addWidget(self.checkbox6)

        # Add buttons to the group box
        self.db_button_layout = QVBoxLayout()

        # Add "Update" button
        self.btn_apply_filter_db = QPushButton("Update")
        self.btn_apply_filter_db.setStyleSheet("background-color: #32612D; color: white; font-size: 16px;")
        self.btn_apply_filter_db.clicked.connect(self.apply_filter_db)
        self.btn_apply_filter_db.setVisible(False)
        self.db_button_layout.addWidget(self.btn_apply_filter_db)

        # Add buttons layout to the group box layout
        self.db_groupbox_layout.addLayout(self.db_button_layout)

        # Add the right and bottom containers to the combined container
        self.right_bottom_layout.addWidget(self.right_container)
        self.right_bottom_layout.addWidget(self.center_container)
        self.right_bottom_layout.addWidget(self.bottom_container)

        # Add the combined container to the main layout
        self.layout.addWidget(self.right_bottom_container)

        # Set the vertical alignment of the right container to the top
        self.right_bottom_layout.setAlignment(Qt.AlignTop)
        self.right_bottom_container.setFixedHeight(971)
        # Set the border for the combined container
        self.right_bottom_container.setStyleSheet("QGroupBox { border: 2px solid gray; }")

        self.activate_checkbox_shortcut = QShortcut(QKeySequence("F6"), self)
        self.activate_checkbox_shortcut.activated.connect(self.activate_selected_checkbox)


    def selecionar_todos(self):
        for row_index in range(self.table_widget.rowCount()):
            checkbox_item = self.table_widget.cellWidget(row_index, 0)
            if isinstance(checkbox_item, QCheckBox) and checkbox_item.isVisible():
                checkbox_item.setChecked(self.select_all_checkbox.isChecked())

    def carregar_ficheiro(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(
            self, "Abrir ficheiro", "", "Ficheiro (*.csv *.xlsx *.xls)")

        if file_path:
            option_dialog = SelecionarOpcao(self)
            if option_dialog.exec_() == QDialog.Accepted:
                selected_option = option_dialog.obter_selecao()
                if selected_option == "conteudos":
                    self.mostrar_opcoes(1)
                elif selected_option == "alteracao_precos":
                    self.mostrar_opcoes(2)
                else:
                    QMessageBox.warning(self, "Aviso", "Nenhuma opção selecionada. Por favor, escolha uma opção.")
                    return

                self.data = self.carregar_dados(file_path)
                if self.data is not None:
                    self.mostrar_dados(self.data)
                    self.reload_filters()
                    self.populate_dropdowns()

                self.centralWidget().setVisible(True)
                self.showMaximized()

    def populate_dropdowns(self):
        if self.data is not None:
            column_names = list(self.data.columns)

            filtered_column_names_1 = [name for name in column_names if 'AQ' in name]

            filtered_column_names_2 = [name for name in column_names if 'AQ' not in name]

            self.combo_box1.clear()
            self.combo_box2.clear()

            self.combo_box1.addItems(filtered_column_names_1)
            self.combo_box2.addItems(filtered_column_names_2)

    def comparar_produtos_selecionados(self):
        selected_rows = []
        for row in range(self.table_widget.rowCount()):
            checkbox_item = self.table_widget.cellWidget(row, 0)
            if checkbox_item.isChecked():
                selected_rows.append(row)

        if not selected_rows:
            QMessageBox.warning(self, "Aviso", "Nenhum produto selecionado. Por favor, selecione pelo menos um produto.")
            return

        # Get the selected column names from dropdowns
        column1 = self.combo_box1.currentText()
        column2 = self.combo_box2.currentText()

        if column1 == column2:
            QMessageBox.warning(self, "Aviso", "Selecione colunas diferentes para comparar.")
            return

        # Listas separadas para cada coluna selecionada
        column1_values = []
        column2_values = []
        descricao_aq_values = []
        ref_aq_values = []

        for row in selected_rows:
            column1_value = self.data.loc[row, column1]
            column1_values.append(column1_value)

            column2_value = self.data.loc[row, column2]
            column2_values.append(column2_value)

            descricao_aq_value = self.data.loc[row, 'Descricao AQ']
            descricao_aq_values.append(descricao_aq_value)

            ref_aq_value = self.data.loc[row, 'Referencia AQ']
            ref_aq_values.append(ref_aq_value)

        # Passando as listas separadas e os nomes das colunas para a janela de comparação
        comparison_dialog = JanelaComparacao(column1_values, column2_values, descricao_aq_values, ref_aq_values, column1, column2)
        comparison_dialog.exec_()
        
    def mostrar_opcoes(self, option):
        if option == 1:
            
            self.btn_apply_filter.setVisible(True)
            self.btn_clear_filters.setVisible(True)
            self.select_all_checkbox.setVisible(True)
            self.btn_apply_filter_db.setVisible(True)
            self.comparar_btn.setVisible(True)

            self.combo_box1.setVisible(True)
            self.combo_box2.setVisible(True)     

            self.checkbox1.setVisible(True)
            self.checkbox2.setVisible(True)
            self.checkbox3.setVisible(True)
            self.checkbox4.setVisible(True)
            self.checkbox5.setVisible(False)
            self.checkbox6.setVisible(False)

        elif option == 2:

            self.btn_apply_filter.setVisible(True)
            self.btn_clear_filters.setVisible(True)
            self.select_all_checkbox.setVisible(True)
            self.btn_apply_filter_db.setVisible(True)
            self.comparar_btn.setVisible(True)

            self.combo_box1.setVisible(True)
            self.combo_box2.setVisible(True)

            self.checkbox1.setVisible(False)
            self.checkbox2.setVisible(False)
            self.checkbox3.setVisible(False)
            self.checkbox4.setVisible(False)
            self.checkbox5.setVisible(True)
            self.checkbox6.setVisible(True)
        else:
            QMessageBox.warning(self, "Aviso", "Opção inválida.")

    def carregar_dados(self, file_path):
        try:
            if file_path.endswith('.csv'):
                with open(file_path, 'r', encoding='utf-8') as csvfile:
                    dialect = csv.Sniffer().sniff(csvfile.read(1024))
                    csvfile.seek(0)
                    delimiter = dialect.delimiter
                    data = pd.read_csv(csvfile, delimiter=delimiter)
            elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                data = pd.read_excel(file_path, engine='openpyxl')
            else:
                raise Exception("Ficheiro não suportado")
            return data

        except FileNotFoundError:
            print("Ficheiro não encontrado.")
            return None
        except Exception as e:  
            print("Ocorreu um erro:", e)
            return None

    # def mostrar_dados(self, data):
    #     mostrar_dados = data.fillna("-")
        
    #     self.table_widget.clear()

    #     num_columns = mostrar_dados.shape[1] + 1

    #     self.table_widget.setRowCount(mostrar_dados.shape[0])
    #     self.table_widget.setColumnCount(num_columns)

    #     header_labels = [""] + list(mostrar_dados.columns)
    #     self.table_widget.setHorizontalHeaderLabels(header_labels)

    #     for i in range(mostrar_dados.shape[0]):
    #         checkbox_item = QCheckBox()
    #         self.table_widget.setCellWidget(i, 0, checkbox_item)

    #         for j in range(mostrar_dados.shape[1]):
    #             value = mostrar_dados.iloc[i, j]
    #             if isinstance(value, float) and value.is_integer():
    #                 value = int(value)
    #             item = QTableWidgetItem(str(value))
    #             self.table_widget.setItem(i, j + 1, item)

    #     self.table_widget.resizeColumnsToContents()
    #     self.table_widget.resize(self.table_widget.horizontalHeader().length() + 50,
    #                             self.table_widget.verticalHeader().length() + 50)


    def mostrar_dados(self, data):
        mostrar_dados = data.fillna("-")
        
        self.table_widget.clear()

        num_columns = mostrar_dados.shape[1] + 1

        self.table_widget.setRowCount(mostrar_dados.shape[0])
        self.table_widget.setColumnCount(num_columns)

        header_labels = [""] + list(mostrar_dados.columns)
        self.table_widget.setHorizontalHeaderLabels(header_labels)

        for i in range(mostrar_dados.shape[0]):
            checkbox_item = QCheckBox()
            self.table_widget.setCellWidget(i, 0, checkbox_item)

            for j in range(mostrar_dados.shape[1]):
                value = mostrar_dados.iloc[i, j]
                if isinstance(value, float) and value.is_integer():
                    value = int(value)
                item = QTableWidgetItem(str(value))
                self.table_widget.setItem(i, j + 1, item)

        # Aplicar a coloração alternada
        self.aplicar_coloracao_alternada()

        self.table_widget.resizeColumnsToContents()
        self.table_widget.resize(self.table_widget.horizontalHeader().length() + 50,
                                self.table_widget.verticalHeader().length() + 50)

    def aplicar_coloracao_alternada(self):
        num_columns = self.table_widget.columnCount()
        for i in range(self.table_widget.rowCount()):
            for j in range(num_columns):
                if j == 0:
                    item = self.table_widget.cellWidget(i, j)
                    if item:
                        if i % 2 == 1:
                            item.setStyleSheet("background-color: rgb(211, 211, 211);")
                        else:
                            item.setStyleSheet("")
                else:
                    item = self.table_widget.item(i, j)
                    if item:
                        if i % 2 == 1:
                            item.setBackground(QColor(211, 211, 211))
                        else:
                            item.setBackground(QColor(255, 255, 255))
    
    def reload_filters(self):
        for widget_layout in self.filter_widgets:
            layout = widget_layout[1]
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

        self.filter_widgets.clear()

        row = 0
        col = 0
        for column_name in self.data.columns:
            if col % 2 == 0:
                row += 1
                col = 0

            filter_layout = QVBoxLayout()
            self.filter_layout.addLayout(filter_layout, row, col)

            combo_box = QComboBox()

            if column_name.lower() in ['ean', 'ref', 'referencia']:
                combo_box.addItems(["igual"])
            else:
                combo_box.addItems(["contém", "igual"])

            combo_box.setCurrentText("igual" if column_name.lower() in ['ean', 'ref', 'referencia'] else "contém")
            filter_layout.addWidget(combo_box)

            line_edit = QLineEdit()
            line_edit.setPlaceholderText(f"Filtrar por {column_name}")
            line_edit.setMaximumWidth(400)
            line_edit.returnPressed.connect(self.apply_filter)
            filter_layout.addWidget(line_edit)

            spacer_item = QSpacerItem(30, 10, QSizePolicy.Expanding, QSizePolicy.Minimum)
            filter_layout.addItem(spacer_item)

            self.filter_widgets.append((column_name, filter_layout))
            col += 1

    def apply_filter(self):
        self.uncheck_all_checkboxes()

        if self.data is None:
            return
        
        active_filters = []
        for column_name, widget_layout in self.filter_widgets:
            line_edit = widget_layout.itemAt(1).widget()
            filter_value = line_edit.text().strip().lower()
            if filter_value:
                combo_box = widget_layout.itemAt(0).widget()
                filter_type = combo_box.currentText() if isinstance(combo_box, QComboBox) else "contém"
                active_filters.append((column_name, filter_value, filter_type))

        for row_index in range(self.table_widget.rowCount()):
            row_visible = True
            for column_name, filter_value, filter_type in active_filters:
                column_index = self.data.columns.get_loc(column_name) + 1
                item = self.table_widget.item(row_index, column_index)
                if item:
                    cell_text = item.text().strip().lower()
                    if filter_type == "contém" and filter_value not in cell_text:
                        row_visible = False
                        break
                    elif filter_type == "igual" and filter_value != cell_text:
                        row_visible = False
                        break
            self.table_widget.setRowHidden(row_index, not row_visible)

    # def apply_filter_db(self):
    #     if self.data is None:
    #         return

    #     records_to_insert = []
    #     print(records_to_insert)
    #     if not (self.checkbox1.isChecked() or self.checkbox2.isChecked() or
    #             self.checkbox3.isChecked() or self.checkbox4.isChecked() or self.checkbox5.isChecked() or self.checkbox6.isChecked()):
    #         QMessageBox.warning(self, "Aviso", "Selecione pelo menos uma opção.")
    #         return
        
    #     for row_index in range(self.table_widget.rowCount()):
    #         checkbox_item = self.table_widget.cellWidget(row_index, 0)
    #         if isinstance(checkbox_item, QCheckBox) and checkbox_item.isChecked():
    #             ean = self.table_widget.item(row_index, 1).text()
    #             referencia = self.table_widget.item(row_index, 2).text()
    #             opcao1 = 1 if self.checkbox1.isChecked() else 0
    #             opcao2 = 1 if self.checkbox2.isChecked() else 0
    #             opcao3 = 1 if self.checkbox3.isChecked() else 0
    #             opcao4 = 1 if self.checkbox4.isChecked() else 0
    #             opcao5 = 1 if self.checkbox5.isChecked() else 0
    #             opcao6 = 1 if self.checkbox6.isChecked() else 0
    #             records_to_insert.append((ean, referencia, opcao1, opcao2, opcao3, opcao4, opcao5, opcao6))

    #     if records_to_insert:
    #         self.guardar_basedados(records_to_insert)
    #         self.uncheck_all_checkboxes()
    #     else:
    #         QMessageBox.warning(self, "Aviso", "Nenhum registo selecionado.")
    def apply_filter_db(self):
        if self.data is None:
            return

        records_to_insert = []

        # Get the column indices for 'EANs AQ' and 'Referencia AQ
        ean_col_index = None
        referencia_col_index = None
        headers = [self.table_widget.horizontalHeaderItem(i).text() for i in range(self.table_widget.columnCount())]

        for i, header in enumerate(headers):
            if header == "EANs AQ":
                ean_col_index = i
            elif header == "Referencia AQ":
                referencia_col_index = i

        if ean_col_index is None or referencia_col_index is None:
            QMessageBox.warning(self, "Erro", "Colunas 'EANs AQ' e/ou 'Referencia AQ' não encontradas.")
            return

        if not (self.checkbox1.isChecked() or self.checkbox2.isChecked() or
                self.checkbox3.isChecked() or self.checkbox4.isChecked() or self.checkbox5.isChecked() or self.checkbox6.isChecked()):
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos uma opção.")
            return

        for row_index in range(self.table_widget.rowCount()):
            checkbox_item = self.table_widget.cellWidget(row_index, 0)
            if isinstance(checkbox_item, QCheckBox) and checkbox_item.isChecked():
                ean_item = self.table_widget.item(row_index, ean_col_index)
                referencia_item = self.table_widget.item(row_index, referencia_col_index)

                # Debugging: Check if the items are None
                if ean_item is None or referencia_item is None:
                    print(f"Row {row_index}: EAN or Referencia is None")
                    continue

                ean = ean_item.text()
                referencia = referencia_item.text()

                # Debugging: Print the values
                print(f"Row {row_index}: EAN = {ean}, Referencia = {referencia}")

                opcao1 = 1 if self.checkbox1.isChecked() else 0
                opcao2 = 1 if self.checkbox2.isChecked() else 0
                opcao3 = 1 if self.checkbox3.isChecked() else 0
                opcao4 = 1 if self.checkbox4.isChecked() else 0
                opcao5 = 1 if self.checkbox5.isChecked() else 0
                opcao6 = 1 if self.checkbox6.isChecked() else 0
                records_to_insert.append((ean, referencia, opcao1, opcao2, opcao3, opcao4, opcao5, opcao6))

        # Debugging: Print the records to insert
        print(f"Records to insert: {records_to_insert}")

        if records_to_insert:
            self.guardar_basedados(records_to_insert)
            self.uncheck_all_checkboxes()
        else:
            QMessageBox.warning(self, "Aviso", "Nenhum registo selecionado.")


    def uncheck_all_checkboxes(self):
        for row_index in range(self.table_widget.rowCount()):
            checkbox_item = self.table_widget.cellWidget(row_index, 0)
            if isinstance(checkbox_item, QCheckBox) and checkbox_item.isChecked():
                checkbox_item.setChecked(False)

        self.checkbox1.setChecked(False)    
        self.checkbox2.setChecked(False)
        self.checkbox3.setChecked(False)
        self.checkbox4.setChecked(False)
        self.checkbox5.setChecked(False)
        self.checkbox6.setChecked(False)

        self.select_all_checkbox.setChecked(False)

    def activate_selected_checkbox(self):
        selected_row = self.table_widget.currentRow()
        if selected_row >= 0:
            checkbox_item = self.table_widget.cellWidget(selected_row, 0)
            if isinstance(checkbox_item, QCheckBox):
                checkbox_item.setChecked(not checkbox_item.isChecked())

    def clear_filters(self):
        self.uncheck_all_checkboxes()
        for row_index in range(self.table_widget.rowCount()):
            self.table_widget.setRowHidden(row_index, False)

        for _, widget_layout in self.filter_widgets:
            line_edit = widget_layout.itemAt(1).widget()
            line_edit.clear()

        self.mostrar_dados(self.data)

    def guardar_basedados(self, records):
        server = 'LAPTOP-1DR8H4K3\SQLEXPRESS'
        database = 'aquariocsv'

        try:
            conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
            cursor = conn.cursor()

            existing_records = {}
            cursor.execute("SELECT EAN, Referencia, opcao1, opcao2, opcao3, opcao4, opcao5, opcao6 FROM aquariocsv.dbo.produto")
            for row in cursor.fetchall():
                existing_records[(row.EAN, row.Referencia)] = (row.opcao1, row.opcao2, row.opcao3, row.opcao4, row.opcao5, row.opcao6)

            for ean, referencia, opcao1, opcao2, opcao3, opcao4, opcao5, opcao6 in records:
                if (ean, referencia) in existing_records:
                    current_opcao1, current_opcao2, current_opcao3, current_opcao4, current_opcao5, current_opcao6 = existing_records[(ean, referencia)]
                    
                    if current_opcao1 == 0:
                        opcao1_to_update = opcao1
                    else:
                        opcao1_to_update = current_opcao1
                    
                    if current_opcao2 == 0:
                        opcao2_to_update = opcao2
                    else:
                        opcao2_to_update = current_opcao2
                    
                    if current_opcao3 == 0:
                        opcao3_to_update = opcao3
                    else:
                        opcao3_to_update = current_opcao3
                    
                    if current_opcao4 == 0:
                        opcao4_to_update = opcao4
                    else:
                        opcao4_to_update = current_opcao4

                    if current_opcao5 == 0:
                        opcao5_to_update = opcao5
                    else:
                        opcao5_to_update = current_opcao5

                    if current_opcao6 == 0:
                        opcao6_to_update = opcao6
                    else:
                        opcao6_to_update = current_opcao6

                    cursor.execute("UPDATE aquariocsv.dbo.produto SET opcao1=?, opcao2=?, opcao3=?, opcao4=?, opcao5=?, opcao6=? WHERE EAN=? AND Referencia=?", (opcao1_to_update, opcao2_to_update, opcao3_to_update, opcao4_to_update, opcao5_to_update, opcao6_to_update, ean, referencia))
                    conn.commit()

            QMessageBox.information(self, "Sucesso", "Dados inseridos/atualizados com sucesso na base de dados.")

        except Exception as e:
            print("Ocorreu um erro ao inserir/atualizar os dados na base de dados:", e)
            QMessageBox.warning(self, "Erro", "Ocorreu um erro ao inserir/atualizar os dados na base de dados.")

def main():
    app = QApplication(sys.argv)
    window = App()
    window.showMaximized()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()