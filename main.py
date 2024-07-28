import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import os


def prettify_xml(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="    ")


def create_xml(language, translations, output_dir):
    # 创建resources根元素
    resources = ET.Element("resources")

    # 添加string元素
    for key, value in translations.items():
        string = ET.SubElement(resources, "string", name=key)
        string.text = value

    # 美化XML字符串
    pretty_xml_as_string = prettify_xml(resources)
    # 保存XML文件
    xml_file = os.path.join(output_dir, f"i18n_output/values-{language}-strings.xml")
    os.makedirs(os.path.dirname(xml_file), exist_ok=True)
    with open(xml_file, 'w', encoding='utf-8') as f:
        f.write(pretty_xml_as_string)
    print(f"Created: {xml_file}")


def convert_xlsx_to_android_strings(xlsx_file):
    # 读取Excel文件
    df = pd.read_excel(xlsx_file)

    # 找到Key列（忽略大小写）
    key_col = None
    for col in df.columns:
        if col.lower() == "key":
            key_col = col
            break

    if key_col is None:
        raise ValueError("Excel文件中没有找到'Key'列")

    # 获取所有语言列（除去Key列）
    languages = [col for col in df.columns if col != key_col]

    # 获取输出目录为输入文件路径的上一级目录
    output_dir = os.path.dirname(xlsx_file)

    for language in languages:
        translations = dict(zip(df[key_col], df[language]))
        create_xml(language, translations, output_dir)


if __name__ == "__main__":
    xlsx_file = input("请输入Excel文件路径: ")
    convert_xlsx_to_android_strings(xlsx_file)
