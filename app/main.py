from app import interface_badges
from libs_dima import load_config
from libs_gleb import save_presentation, merge_all_classes
from generator import generate
from get_excel_list import fetch_dict_classes
import os

#Функция для получения списка участников по классу
def list_for_classname():
    config = load_config()
    excel_dict = fetch_dict_classes(config["filepath_xls"])
    if config['all']:
        return merge_all_classes(excel_dict)
    return excel_dict[config['class']]


interface_badges()
config = load_config()

if config['submitted']:
    if config['all']:
        result_path = r"result\result_all.pptx"
    else:
        result_path = rf"result\result_{config['class']}.pptx"
    root_slide = (config["slide"])-1
    badge_class_list = list_for_classname()
    filepath_pptx = config['filepath_pptx']

    result = generate(filepath_pptx, badge_class_list, root_slide)
    save_presentation(result, result_path)

    os.startfile(result_path)