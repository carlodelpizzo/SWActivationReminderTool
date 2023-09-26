from main import main
import os

appdata_path = os.environ['APPDATA']
working_folder = 'SWReminderTool'
working_dir = f'{appdata_path}/{working_folder}'

if __name__ == '__main__':
    if not os.path.isdir(working_dir):
        os.mkdir(working_dir)
    if not os.path.isfile(f'{working_dir}/running'):
        with open(f'{working_dir}/running', 'w', errors='ignore') as out_file:
            out_file.writelines('')
    elif not os.path.isfile(f'{working_dir}/checkin'):
        with open(f'{working_dir}/checkin', 'w', errors='ignore') as out_file:
            out_file.writelines('')
        waiting = True
    else:
        checkin = True

    main()
