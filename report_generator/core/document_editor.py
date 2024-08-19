class DocumentEditor:
    def __init__(self, doc):
        self.doc = doc

    def replace_report_text(self, replacements):
        # 替换段落中的占位符
        for paragraph in self.doc.paragraphs:
            runs = paragraph.runs
            for i, run in enumerate(runs):
                if run.text == '#':
                    counter = i  # 记录起始位置
                    tmp = '#'    # tmp开始存储
                    while tmp not in list(replacements.keys()):
                        counter += 1
                        tmp += runs[counter].text
                        runs[counter].clear()
                    runs[i].text = runs[i].text.replace(runs[i].text, replacements[tmp])

        # 替换表格中的占位符
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        runs = paragraph.runs
                        full_text = ''.join(run.text for run in runs)

                        for key, value in replacements.items():
                            if key in full_text:
                                remaining_text = full_text
                                for run in runs:
                                    if key in remaining_text:
                                        start_index = remaining_text.index(key)
                                        end_index = start_index + len(key)

                                        pre_key_text = remaining_text[:start_index]
                                        post_key_text = remaining_text[end_index:]

                                        run.text = pre_key_text + value
                                        remaining_text = post_key_text
                                    else:
                                        run.text = remaining_text
                                        remaining_text = ''
