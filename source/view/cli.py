import argparse
from source.controller.controller import Controller


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Data Analyser', description='')
    parser.add_argument('load_path', help='')
    parser.add_argument('save_path', help='')
    parser.add_argument('-r', '--rule', default='settings/rule/data.json', help='')
    parser.add_argument('-v', '--validation', default='settings/validation/data.json', help='')
    args = parser.parse_args()

    Controller(args.load_path, args.save_path).column_fill(args.rule)
    Controller(args.save_path, args.save_path).validation_add(args.validation)
