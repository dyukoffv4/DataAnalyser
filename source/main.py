import argparse
import json
import sys
import os


def standard():
    parser = argparse.ArgumentParser(prog='Data Analyser', description='Fills in Excel column data based on established rules')
    parser.add_argument('-r', '--rule', default='settings/rule/data.json')
    parser.add_argument('-d', '--data_validation', nargs='?', const='settings/validation/data.json')
    parser.add_argument('work_path')
    parser.add_argument('save_path')
    args = parser.parse_args()

    if args.data_validation is not None:
        print('-d is not able yet.')
        # Todo: make it!

    # Todo: open and process .json.
    # with open(args.rule, 'r') as f:
    #     data = json.load(f)
    #     print(data)

    if os.path.isdir(args.work_path):
        for i in os.listdir(args.work_path):
            if i.split('.')[-1] == 'xlsx':
                print(i)  # Todo: insert worker
    else:
        print(args.work_path)  # Todo: insert worker


def create_rule():
    print('create is not able yet.')
    # Todo: make it!


if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == 'create':
        if sys.argv[2] == 'rule':
            exit(create_rule())
        raise RuntimeError('Have not such creator')
    exit(standard())
