
import warnings
import tldextract
import pandas as pd
from docx import Document
from datetime import datetime
from core.document_image_processor import DocumentImageProcessor
from core.report_generator import ReportGenerator
from core.excel_data_reader import ExcelDataReader
from core.document_editor import DocumentEditor
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QListView, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QFormLayout, QMessageBox, QScrollArea
warnings.filterwarnings("ignore", category=DeprecationWarning)

class MainWindow(QWidget):
    def __init__(self, push_config):
        super().__init__()
        
        # 从 YAML 文件中获取默认值
        self.push_config = push_config

        # 保存所有的漏洞复现描述部分
        self.vuln_sections = []

        # 创建 ExcelDataReader 对象，并进行处理
        self.excel_data_reader = ExcelDataReader()

        # 从Excel文件中读取ICP信息
        self.Icp_infos = self.excel_data_reader.read_Icp_from_excel(self.push_config["icp_info_file"])

        # 从Excel文件中读取漏洞信息
        self.vulnerability_names, self.vulnerabilities = self.excel_data_reader.read_vulnerabilities_from_excel(self.push_config["vulnerabilities_file"])

        # 设置窗口标题和图标     
        self.setWindowTitle(f'风险隐患报告生成器 - {self.push_config["version"]}')
        self.setWindowIcon(QIcon(self.push_config["icon_path"]))

        self.setFixedSize(620, 700)
        self.init_ui()  # 初始化UI界面

    def init_ui(self):

        '''设置 GUI 组件的初始化代码'''
        self.labels = ["隐患编号:", "隐患名称:", "隐患URL:", "隐患类型:", "隐患级别:",
                       "预警级别:", "归属地市:", "单位类型:", "所属行业:", "单位名称:",
                       "网站名称:", "网站域名:", "网站IP:", "备案号:", "发现时间:",
                       "漏洞描述:", "漏洞危害:", "修复建议:", "证据截图:", "工信备案截图:", 
                       "备注:"]
        

        self.text_edits = [QLineEdit(self) for _ in range(15)]

        # 创建文本框用于隐患编号
        self.vulnerability_id_text_edit = self.text_edits[0]
        # 设置文本框的初始文本为生成的隐患编号
        self.vulnerability_id_text_edit.setText(self.generate_vulnerability_id())
        # 创建漏洞类型下拉框
        self.vulName_box = QComboBox(self)
        self.vulName_box.addItems(self.vulnerability_names)
        self.setup_combobox_style(self.vulName_box, 200)

        # 创建漏洞等级下拉框
        self.hazardLevel_box = QComboBox(self)
        self.hazardLevel_box.addItems(['高危', '中危', '低危'])
        self.setup_combobox_style(self.hazardLevel_box, 70)

        # 创建文本框用于预警级别
        self.alert_level_text_edit = self.text_edits[1]
        self.alert_level_text_edit.setReadOnly(True)  # 只读
        # 当hazardLevel_box值改变时调用update_alert_level方法
        self.hazardLevel_box.currentIndexChanged.connect(self.update_alert_level)
        # 初始化预警级别
        self.update_alert_level()

        # 创建单位类型下拉框
        self.unitType_box = QComboBox(self)
        self.unitType_box.addItems(self.push_config["unitType"])
        self.setup_combobox_style(self.unitType_box, 100)

        # 创建所属行业下拉框
        self.industry_box = QComboBox(self)
        self.industry_box.addItems(self.push_config["industry"])
        self.setup_combobox_style(self.industry_box, 100)

        # 创建文本框用于发现时间
        self.discovery_date_edit = self.text_edits[13]
        # 设置文本框的初始文本为当前日期
        self.discovery_date_edit.setText(datetime.now().strftime('%Y.%m.%d'))

        # 创建用于显示工信备案截图的标签和按钮
        self.image_label_asset = QLabel(self)
        self.paste_button_asset = QPushButton('点击读取截图', self)
        self.delete_button_asset = QPushButton('删除图片', self)
        self.paste_button_asset.clicked.connect(self.paste_asset_image)
        self.delete_button_asset.clicked.connect(self.delete_asset_image)

        # 添加按钮用于在界面上添加新的漏洞复现描述和漏洞证明图片的功能
        self.add_vuln_button = QPushButton('添加证明', self)
        self.generate_button = QPushButton('生成报告', self)
        self.reset_button = QPushButton('一键重置', self)
        self.clear_all_button = QPushButton('一键清除', self)
        self.add_vuln_button.clicked.connect(self.add_vulnerability_section)
        self.generate_button.clicked.connect(self.generate_report)
        self.reset_button.clicked.connect(self.reset_all)
        self.clear_all_button.clicked.connect(self.clear_all_sections)

        '''设置 GUI 组件表单布局'''
        self.form_layout = QFormLayout()
        self.setup_formlayout()
        self.setup_main_layout()

    '''设置下拉框样式'''
    def setup_combobox_style(self, combobox, width):
        combobox.setFixedSize(width, 20)
        combobox.setView(QListView())   ##todo 下拉框样式
        combobox.setStyleSheet("QComboBox QAbstractItemView {font-size:14px;}"     # 下拉文字大小
                               "QComboBox QAbstractItemView::item {height:30px;padding-left:20px;}"  # 下拉文字宽高
                               "QScrollBar:vertical {border:2px solid grey;width:20px;}")    # 下拉侧边栏宽高

    def setup_formlayout(self):
        # 添加用于隐患编号的文本框到布局
        self.form_layout.addRow(QLabel(self.labels[0]), self.vulnerability_id_text_edit)

        # 创建一个水平布局用于放置漏洞类型和漏洞等级
        h_layout = QHBoxLayout()
        # 添加漏洞类型下拉框到表单布局
        h_layout.addWidget(self.vulName_box)
        # 添加漏洞等级下拉框到表单布局
        h_layout.addWidget(QLabel(self.labels[4]))
        h_layout.addWidget(self.hazardLevel_box)
        # 添加自动更新预警级别到表单布局
        h_layout.addWidget(QLabel(self.labels[5]))
        h_layout.addWidget(self.alert_level_text_edit)
        self.form_layout.addRow(QLabel(self.labels[3]), h_layout)

        # 添加隐患URL到表单布局
        self.form_layout.addRow(QLabel(self.labels[2]), self.text_edits[5])
        self.text_edits[5].textChanged.connect(self.update_get_domain)
       
        # 创建一个水平布局用于放置域名信息
        Website_Name_layout = QHBoxLayout()
        # 添加网站名称到表单布局
        Website_Name_layout.addWidget(self.text_edits[6])
        # 添加网站IP到表单布局
        Website_Name_layout.addWidget(QLabel(self.labels[12]))
        Website_Name_layout.addWidget(self.text_edits[8])
        self.form_layout.addRow(QLabel(self.labels[10]), Website_Name_layout)

        # 创建一个水平布局用于放置域名信息
        domain_layout = QHBoxLayout()
        # 添加网站域名到表单布局
        domain_layout.addWidget(self.text_edits[7])
        # 添加工信备案号到表单布局
        domain_layout.addWidget(QLabel(self.labels[13]))
        domain_layout.addWidget(self.text_edits[9])
        self.form_layout.addRow(QLabel(self.labels[11]), domain_layout)
        # 添加单位名称到表单布局
        self.form_layout.addRow(QLabel(self.labels[9]), self.text_edits[3])

        # 在init_ui方法中，为网站域名添加信号，根据域名自动从文件中提取备案信息
        self.text_edits[7].textChanged.connect(self.update_icp_info)

        # 在init_ui方法中，为单位名称、网站名称和隐患类型添加信号，根据其中变化自动调整其他参数数据
        self.text_edits[3].textChanged.connect(self.update_hazard_name)
        self.text_edits[6].textChanged.connect(self.update_hazard_name)
        self.vulName_box.currentIndexChanged.connect(self.update_hazard_name)

        # 创建一个水平布局用于放置公司信息
        unit_layout = QHBoxLayout()
        # 添加归属城市到表单布局
        self.text_edits[4].setText(self.push_config["city"])  # 设置默认值
        unit_layout.addWidget(self.text_edits[4])
        # 添加单位类型到表单布局
        unit_layout.addWidget(QLabel(self.labels[7]))
        unit_layout.addWidget(self.unitType_box)
        # 添加所属行业到表单布局
        unit_layout.addWidget(QLabel(self.labels[8]))
        unit_layout.addWidget(self.industry_box)
        self.form_layout.addRow(QLabel(self.labels[6]), unit_layout)

        # 添加隐患名称到表单布局
        self.form_layout.addRow(QLabel(self.labels[1]), self.text_edits[2])

        # 添加发现时间到表单布局
        self.form_layout.addRow(QLabel(self.labels[14]), self.discovery_date_edit)

        # 添加漏洞描述到表单布局
        self.form_layout.addRow(QLabel(self.labels[15]), self.text_edits[10])
        self.text_edits[10].setFixedHeight(60)  # 漏洞描述可能较长，增加文本框高度
        # 添加漏洞危害到表单布局
        self.form_layout.addRow(QLabel(self.labels[16]), self.text_edits[11])
        self.text_edits[11].setFixedHeight(60)  # 漏洞危害可能较长，增加文本框高度
        self.text_edits[11].textChanged.connect(self.update_hazard_name)
        # 添加整改建议到表单布局
        self.form_layout.addRow(QLabel(self.labels[17]), self.text_edits[12])
        self.text_edits[12].setFixedHeight(60)  # 整改建议可能较长，增加文本框高度

        # 添加备注到表单布局
        self.form_layout.addRow(QLabel(self.labels[-1]), self.text_edits[14])

        # 创建按钮布局用于放置生成报告
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.generate_button)
        button_layout.addWidget(self.reset_button)  # 一键重置
        self.form_layout.addRow(button_layout)

        # 添加新的漏洞复现描述和图片按钮
        vuln_button_layout = QHBoxLayout()
        vuln_button_layout.addWidget(self.add_vuln_button)
        vuln_button_layout.addWidget(self.clear_all_button)
        self.form_layout.addRow(vuln_button_layout)

        # 添加工信备案截图到表单布局
        asset_layout = QHBoxLayout()
        asset_layout.addWidget(self.image_label_asset)
        asset_layout.addWidget(self.paste_button_asset)
        asset_layout.addWidget(self.delete_button_asset)
        self.form_layout.addRow(QLabel(self.labels[19]), asset_layout)

        # 设置默认值
        self.update_hazard_name()
        self.add_vulnerability_section()

    '''把表单布局添加到主布局中'''
    def setup_main_layout(self):
        
        # 创建一个垂直布局，用于管理其他小部件和布局
        v_layout = QVBoxLayout()

        # 创建一个滚动区域，用于容纳可能超出屏幕显示范围的内容
        v_scroll = QScrollArea()

        # 将表单布局添加到垂直布局中
        v_layout.addLayout(self.form_layout)

        # 创建一个QWidget作为滚动区域的子部件
        widget = QWidget()

        # 将垂直布局设置为widget的布局
        widget.setLayout(v_layout)

        # 将widget设置为滚动区域的子部件
        v_scroll.setWidget(widget)

        # 设置滚动区域可以自动调整大小以适应其内容
        v_scroll.setWidgetResizable(True)

        # 指定滚动区域的宽度
        v_scroll.setFixedWidth(600)  # 或者使用 setMinimumWidth 根据需要
        # 创建主布局，用于管理整个窗口的内容
        main_layout = QVBoxLayout()

        # 将滚动区域添加到主布局中
        main_layout.addWidget(v_scroll)

        # 将主布局设置为QMainWindow的布局
        self.setLayout(main_layout)

        # 显示窗口
        self.show()

    '''根据系统时间生成隐患编号'''
    def generate_vulnerability_id(self):
        current_time = datetime.now().strftime('%Y-%m-%d-%H%M%S')
        return f"HN-XX-XX-{current_time}"
    
    '''重置所有数据'''
    def reset_all(self):
        self.vulnerability_id_text_edit.setText(self.generate_vulnerability_id())
        self.vulName_box.setCurrentIndex(0)
        self.hazardLevel_box.setCurrentIndex(0)
        # self.alert_level_text_edit.clear()    # 等级不需要清除
        self.unitType_box.setCurrentIndex(0)
        self.industry_box.setCurrentIndex(0)
        self.text_edits[2].clear()
        self.text_edits[3].clear()
        # self.text_edits[4].clear()    # 城市不需要清除
        self.text_edits[5].clear()
        self.text_edits[6].clear()
        self.text_edits[7].clear()
        self.text_edits[8].clear()
        self.text_edits[9].clear()
        self.text_edits[10].clear()
        self.text_edits[11].clear()
        self.text_edits[12].clear()
        self.text_edits[13].setText(datetime.now().strftime('%Y.%m.%d'))
        self.text_edits[14].clear()
        self.delete_asset_image()
        self.clear_all_sections()

    '''仅清除漏洞复现数据'''
    def clear_all_sections(self):
        for section in self.vuln_sections:
            layout, edit, image_label = section
            layout.deleteLater()
            edit.deleteLater()
            image_label.deleteLater()
            self.form_layout.removeRow(layout)
        self.vuln_sections.clear()

        # 设置默认值
        self.vulnerability_id_text_edit.setText(self.generate_vulnerability_id())
        self.update_hazard_name()
        self.add_vulnerability_section()
    
    '''提取隐患url的根域名'''
    def update_get_domain(self):
        url = self.text_edits[5].text().strip()
        domain = tldextract.extract(url).registered_domain
        self.text_edits[7].setText(domain)  # 设置网站域名

    '''根据域名自动识别ICP备案信息'''
    def update_icp_info(self):
        domain = self.text_edits[7].text().strip()

        # 根据根域名获取单位名称和备案号
        unit_name, service_licence = self.excel_data_reader.get_Icp_info(domain)

        self.text_edits[3].setText(unit_name)
        self.text_edits[9].setText(service_licence)

    '''添加一个槽函数用于更新隐患名称的值'''
    def update_hazard_name(self):
        unit_name = self.text_edits[3].text().strip()
        website_name = self.text_edits[6].text().strip()
        Vulnerability_Hazard = self.text_edits[11].text().strip()
        hazard_type = self.vulName_box.currentText()
        hazard_name = f"{unit_name}{website_name}存在{hazard_type}漏洞隐患"
        self.text_edits[2].setText(hazard_name)  # 设置隐患名称

        # 根据漏洞名称获取漏洞描述和加固建议
        description, solution = self.excel_data_reader.get_vulnerability_info(hazard_type)
        
        # 检查并打印出哪些变量是 NaN, 也就是列表内存在空值, 如果为NaN将其替换为空字符串
        description = "" if pd.isna(description) else description
        solution = "" if pd.isna(solution) else solution

        # 设置漏洞描述
        if description:
            description_text = f"{description}{Vulnerability_Hazard}" if len(Vulnerability_Hazard) > 0 else description
        else:
            description_text = f"{hazard_name}{Vulnerability_Hazard}" if len(Vulnerability_Hazard) > 0 else hazard_name

        self.text_edits[10].setText(description_text)  # 设置漏洞描述
        self.text_edits[12].setText(solution)  # 设置整改建议

    '''更新预警级别'''
    def update_alert_level(self):
        hazard_level = self.hazardLevel_box.currentText()
        alert_level_map = {
            '高危': '3级',
            '中危': '4级',
            '低危': '4级'
        }
        alert_level = alert_level_map.get(hazard_level, '')
        self.alert_level_text_edit.setText(alert_level)

    def add_vulnerability_section(self):
        # 创建一个新的水平布局用于漏洞复现描述和相关操作按钮
        new_vuln_layout = QHBoxLayout()

        # 创建编辑框、标签和按钮
        new_vuln_edit = QLineEdit(self)
        new_vuln_image_label = QLabel(self)
        new_paste_button = QPushButton('点击读取截图', self)
        new_paste_button.clicked.connect(lambda: self.paste_new_vuln_image(new_vuln_image_label))
        new_delete_button = QPushButton('删除图片', self)
        new_delete_button.clicked.connect(lambda: self.delete_new_vuln_image(new_vuln_image_label))
        delete_section_button = QPushButton('删除该段', self)
        delete_section_button.clicked.connect(lambda: self.delete_vulnerability_section(new_vuln_layout, new_vuln_edit, new_vuln_image_label))

        # 将部件添加到新的水平布局中
        new_vuln_layout.addWidget(QLabel("漏洞复现描述:"))
        new_vuln_layout.addWidget(new_vuln_edit)
        new_vuln_layout.addWidget(new_vuln_image_label)
        new_vuln_layout.addWidget(new_paste_button)
        new_vuln_layout.addWidget(new_delete_button)
        new_vuln_layout.addWidget(delete_section_button)
        self.form_layout.addRow(new_vuln_layout)

        # 保存漏洞复现描述和图片路径
        self.vuln_sections.append((new_vuln_layout, new_vuln_edit, new_vuln_image_label))

    '''监控剪贴板'''
    def get_screenshot_from_clipboard(self):
        clipboard = QApplication.clipboard()
        mime_data = clipboard.mimeData()
        # 检查剪贴板数据是否为图片类型
        if mime_data.hasImage():
            # 获取 QImage 对象而不是原始数据
            image = clipboard.image()
            return image
        else:
            QMessageBox.warning(self, '错误', '剪贴板中没有图片！')
            return None
    
    '''处理备案截图'''
    def paste_asset_image(self):
        """粘贴图像到 QLabel 并保存图像路径"""
        screenshot = self.get_screenshot_from_clipboard()
        if screenshot:
            self.asset_image = screenshot
            # pyqt6: 在 GUI 中显示缩放后的图片
            self.image_label_asset.setPixmap(QPixmap.fromImage(screenshot).scaled(50, 50, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            # 保存原始大小图片的引用
            self.image_label_asset.original_pixmap = QPixmap.fromImage(screenshot)

    '''处理漏洞复现截图'''
    def paste_new_vuln_image(self, image_label):
        screenshot = self.get_screenshot_from_clipboard()
        if screenshot:
            # pyqt6: 在 GUI 中显示缩放后的图片
            image_label.setPixmap(QPixmap.fromImage(screenshot).scaled(50, 50, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            # 保存原始大小图片的引用
            image_label.original_pixmap = QPixmap.fromImage(screenshot)
    
    '''删除备案图片'''
    def delete_asset_image(self):
        self.image_label_asset.clear()

    '''删除复现图片'''
    def delete_new_vuln_image(self, image_label):
        image_label.clear()
        
    '''删除该段'''
    def delete_vulnerability_section(self, layout, edit, label):
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        self.form_layout.removeRow(layout)
        self.vuln_sections.remove((layout, edit, label))

    '''生成报告'''
    def generate_report(self):
        # 加载模板文件
        self.doc = Document(self.push_config["template_path"])

        # 创建 DocumentEditor 对象，并进行处理
        self.editor = DocumentEditor(self.doc)
        
        # 创建 ScreenshotHandler 实例并调用相应的函数
        # self.handler = ScreenshotHandler(self.doc, self.vuln_sections)

        # 创建 DocumentImageProcessor 对象，并进行处理
        self.image_processor = DocumentImageProcessor(self.doc, self.vuln_sections)

        # 创建 ReportGenerator 对象，并进行处理
        self.report_generator = ReportGenerator(self.doc, self.push_config["output_filepath"], self.push_config["supplierName"])

        # 创建一个字典，包含所有需要替换的字段
        replacements = {
            '#reportId#': self.text_edits[0].text().strip(),
            '#reportName#': self.text_edits[2].text().strip(),
            '#target#': self.text_edits[5].text().strip(),
            '#vulName#': self.vulName_box.currentText(),
            '#hazardLevel#': self.hazardLevel_box.currentText(),
            '#warningLevel#': self.alert_level_text_edit.text().strip(),
            '#city#': self.text_edits[4].text().strip(),
            '#unitType#': self.unitType_box.currentText(),
            '#industry#': self.industry_box.currentText(),
            '#customerCompanyName#': self.text_edits[3].text().strip(),
            '#websitename#': self.text_edits[6].text().strip(),
            '#domain#': self.text_edits[7].text().strip(),
            '#ipaddress#': self.text_edits[8].text().strip(),
            '#caseNumber#': self.text_edits[9].text().strip(),
            '#reportTime#': self.discovery_date_edit.text().strip(),
            '#problemDescription#': self.text_edits[10].text().strip(),
            '#vul_modify_repair#': self.text_edits[12].text().strip(),
            '#remark#': self.text_edits[14].text().strip(),
        }
        self.editor.replace_report_text(replacements)

        # 添加工信备案截图
        if hasattr(self, 'asset_image') and self.asset_image:
            asset_path = self.image_processor.save_image_temporarily(self.asset_image)
            self.image_processor.text_with_image("#screenshotoffiling#", asset_path)


        # 处理单个或多个漏洞复现描述和图片
        self.image_processor.process_vuln_sections()

        # 保存日志及文件
        report_file_path = self.report_generator.log_save(replacements)

        # 显示一个消息框通知用户报告已生成
        QMessageBox.information(None, '报告生成', f'报告已生成: {report_file_path}')
