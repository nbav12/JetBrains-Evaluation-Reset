import ctypes
import shutil
import time
import os
import sys
import winreg
import xml.etree.ElementTree as Et
from typing import NoReturn

SUCCESS = '\033[92m'
INFO = '\033[93m'
FAIL = '\033[91m'
END = '\033[0m'

APP_DATA_PATH = os.environ.get('APPDATA')
HOME_PATH = os.path.join(os.environ.get('HOMEDRIVE'), os.environ.get('HOMEPATH'))
REG_PATH = r'SOFTWARE\JavaSoft\Prefs\jetbrains'


def enable_vt_100():
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)


def is_valid_choose(choose: str, min_value: int, max_value: int) -> bool:
    try:
        user_choose = int(choose)
        return min_value <= user_choose <= max_value
    except ValueError:
        return False


def get_product_name() -> str:
    products_options = {
        1: {'value': 'webstorm', 'desc': 'Webstorm'},
        2: {'value': 'pycharm', 'desc': 'PyCharm'}
    }
    print('# Choose Product')
    for key, val in products_options.items():
        print(f'{str(key)}. {val.get("desc")}')

    choose = input('Your product: ')
    while not is_valid_choose(choose, 1, len(products_options)):
        choose = input('Your product: ')

    return products_options.get(int(choose)).get('value')


def close() -> NoReturn:
    os.system('pause')
    sys.exit(0)


def find_products_dirs(product_name: str) -> list:
    products_dirs = []

    # Trying get with 2020.1 and above versions
    jetbrains_dir = os.path.join(APP_DATA_PATH, 'Jetbrains')

    if os.path.exists(jetbrains_dir):
        for root, dirs, files in os.walk(jetbrains_dir):
            for directory in dirs:
                # Ignore PyCharm Commercial Edition (CE)
                if directory.lower().startswith('pycharmce'):
                    continue
                elif directory.lower().startswith(product_name):
                    products_dirs.append(os.path.join(jetbrains_dir, directory))
            break  # Disable recursive of os.walk()

    # Trying get with 2019.3.x and below versions
    for root, dirs, files in os.walk(HOME_PATH):
        for directory in dirs:
            # Ignore PyCharm Commercial Edition (CE)
            if 'pycharmce' in directory.lower():
                continue
            elif product_name in directory.lower():
                product_dir = os.path.join(HOME_PATH, directory)
                if 'config' in os.listdir(product_dir):
                    products_dirs.append(os.path.join(product_dir, 'config'))
        break

    return products_dirs


def choose_product_dir_manual() -> str:
    from tkinter import filedialog, Tk

    root = Tk()
    root.withdraw()

    product_dir = filedialog.askdirectory(title='Choose your product config folder')

    if not product_dir:
        print(f'{INFO}[!] No folder selected. Exiting{END}')
        close()

    return product_dir


def choose_specific_dirs(products_dirs: list) -> list:
    print(f'{INFO}[!] {len(products_dirs)} dirs have been found.{END}')
    print('Which of them do you want to reset?')

    for i, directory in enumerate(products_dirs):
        print(f'\t{i + 1} - {directory}')
    print(f'\t{i + 2} - All')

    user_choose = input('Your choose: ')
    while not is_valid_choose(user_choose, 1, len(products_dirs) + 1):
        user_choose = input('Your choose: ')

    if int(user_choose) <= len(products_dirs):
        return [products_dirs[int(user_choose) - 1]]
    return products_dirs


def find_dirs(dir_name: str, products_dirs: list) -> list:
    return [os.path.join(directory, dir_name) for directory in products_dirs if dir_name in os.listdir(directory)]


def remove_eval_dirs(eval_dirs: list) -> None:
    for directory in eval_dirs:
        shutil.rmtree(directory)
        print(f'{SUCCESS}[+] evl dir {directory} has been removed{END}')


def handle_eval(products_dirs: list) -> None:
    eval_dirs = find_dirs('eval', products_dirs)

    if not eval_dirs:
        return print(f'{INFO}[!] Can not find any eval dir. Skipping{END}')

    remove_eval_dirs(eval_dirs)


def remove_xml_elements(options_dirs: list) -> None:
    xml_files = ('options.xml', 'other.xml')

    for options_dir in options_dirs:
        for file in xml_files:
            xml_path = os.path.join(options_dir, file)

            try:
                tree = Et.parse(xml_path)
                print(f'{SUCCESS}[+] XML file: {xml_path} found. Handling{END}')
            except FileNotFoundError:
                continue

            root = tree.getroot()

            # Filter the component element of the properties by checking "name" attribute
            properties_component = \
                tuple(filter(lambda el: el.get('name') == 'PropertiesComponent', root.findall('component')))[0]

            for element in properties_component[:]:
                if element.get('name').startswith('evl'):
                    properties_component.remove(element)
                    print(f"{SUCCESS}[+] {element.get('name')} element has been removed{END}")

            # Saving xml file and break, because the desired file has been found
            tree.write(xml_path)
            break


def handle_xml(products_dirs: list) -> None:
    options_dirs = find_dirs('options', products_dirs)

    if not options_dirs:
        return print(f'{INFO}[!] Can not find any options dir. Skipping{END}')

    remove_xml_elements(options_dirs)


def delete_sub_keys(key):
    """
    Because it is impossible to delete key with sub keys, there is a need to delete the sub keys recursively.
    After that, delete the root key itself.
    :param key: The desired root key to delete its sub keys
    :return: None
    """
    sub_keys_amount = winreg.QueryInfoKey(key)[0]

    for sub_key_index in range(sub_keys_amount):
        sub_key_name = winreg.EnumKey(key, 0)

        try:
            winreg.DeleteKey(key, sub_key_name)
            print(f'{SUCCESS}[+] {sub_key_name} key has been removed{END}')
        except PermissionError:
            delete_sub_keys(winreg.OpenKey(key, sub_key_name))

    winreg.DeleteKey(key, '')


def handle_reg(product_name):
    reg_path = REG_PATH + rf'\{product_name}'

    # Running over all sub keys of the root key and delete their subs and themself
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as product_key:
            sub_keys_names = [winreg.EnumKey(product_key, sub_key_index)
                              for sub_key_index in range(winreg.QueryInfoKey(product_key)[0])]

            for sub_key_name in sub_keys_names:
                with winreg.OpenKey(product_key, sub_key_name) as sub_key:
                    delete_sub_keys(sub_key)
                    print(f'{SUCCESS}[+] {sub_key_name} key has been removed{END}')
    except FileNotFoundError:
        print(f'{INFO}[!] {reg_path} registry can not be found. Skipping{END}')
    except PermissionError:
        print(f'{FAIL}[-] There is not permission for {reg_path} registry{END}')


def main():
    product_name = get_product_name()
    product_dirs = find_products_dirs(product_name)

    if not product_dirs:
        print(f'{INFO}[!] The {product_name} config dir can not be found. Please choose it manually{END}')
        time.sleep(2)
        product_dirs.append(choose_product_dir_manual())
    elif len(product_dirs) > 1:
        product_dirs = choose_specific_dirs(product_dirs)

    print(f'\n{INFO}[!] Eval dir{END}')
    handle_eval(product_dirs)
    print(f'\n{INFO}[!] XML file{END}')
    handle_xml(product_dirs)
    print(f'\n{INFO}[!] Registry{END}')
    handle_reg(product_name)
    print(f'\n{INFO}[!] Done!{END}')
    close()


if __name__ == '__main__':
    enable_vt_100()
    main()
