import tkinter as tk
from pywinauto import application, findwindows, findbestmatch, controls


def auto_activate():
    def first_click(app):
        dlg = app.top_window()
        try:
            if 'Thank you for installing SOLIDWORKS' in dlg.Static2.texts()[0]:
                dlg['&Next >'].click()
            else:
                first_click(app)
        except findbestmatch.MatchError:
            finish_click(app)

    def second_click(app):
        dlg = app.top_window()
        try:
            if 'To activate your SOLIDWORKS product you must request a license key' in dlg.Edit1.texts()[0]:
                dlg['Select All'].click()
                dlg['&Next >'].click()
            else:
                second_click(app)
        except findbestmatch.MatchError:
            second_click(app)

    def finish_click(app):
        dlg = app.top_window()
        try:
            if dlg.GroupBox1.texts()[0] == 'Result':
                dlg['Finish'].click()
            else:
                finish_click(app)
        except findbestmatch.MatchError:
            finish_click(app)

    try:
        findwindows.find_window(title_re='SOLIDWORKS Professional*')
        return True
    except findwindows.WindowNotFoundError:
        pass

    try:
        window_handle = findwindows.find_window(title='SOLIDWORKS Product Activation')
        sw_app = application.Application().connect(handle=window_handle)
        first_click(sw_app)
        second_click(sw_app)
        finish_click(sw_app)
        return True
    except findwindows.WindowNotFoundError:
        return False


def monitor():
    try:
        findwindows.find_window(title_re='SOLIDWORKS Professional*')
        activation_state = True
        sw_open = True
    except findwindows.WindowNotFoundError:
        return monitor()
    monitor_loop = True
    while monitor_loop:
        try:
            findwindows.find_window(title_re='SOLIDWORKS Professional*')
        except findwindows.WindowNotFoundError:
            sw_open = False
            break
        waiting_for_deactivation = False
        sw_app = None
        try:
            window_handle = findwindows.find_window(title='SOLIDWORKS Product Activation')
            sw_app = application.Application().connect(handle=window_handle)
            sw_dlg = sw_app.top_window()
            try:
                if 'deactivate' in sw_dlg.Edit1.texts()[0]:
                    sw_dlg['Select All'].click()
                    sw_dlg['&Next >'].click()
                    waiting_for_deactivation = True
            except (findbestmatch.MatchError, controls.hwndwrapper.InvalidWindowHandle):
                pass
        except findwindows.WindowNotFoundError:
            pass
        while waiting_for_deactivation:
            sw_dlg = sw_app.top_window()
            try:
                if sw_dlg.GroupBox1.texts()[0] == 'Result':
                    sw_dlg['Finish'].click()
                    sw_open = False
                    activation_state = False
                    monitor_loop = False
                    waiting_for_deactivation = False
            except findbestmatch.MatchError:
                pass
    if not sw_open and activation_state:
        def null_func():
            pass

        def wait_for_reopen(win):
            waiting_for_reopen = True
            try:
                findwindows.find_window(title_re='SOLIDWORKS Professional*')
                win.destroy()
                waiting_for_reopen = False
            except findwindows.WindowNotFoundError:
                pass

            if waiting_for_reopen:
                win.after(1000, lambda: wait_for_reopen(win))

        root = tk.Tk()
        root.geometry('500x175')
        root.title('Deactivation Reminder')
        root.resizable(False, False)
        root.protocol('WM_DELETE_WINDOW', null_func)

        reminder_label = tk.Label(root,
                                  text='Don\'t forget to deactivate SW\n\nre-open SolidWorks\nto close this window',
                                  font=('Arial', 15))
        reminder_label.pack()

        root.after(1000, lambda: wait_for_reopen(root))
        root.mainloop()


def main():
    try:
        findwindows.find_window(title_re='SOLIDWORKS Professional*')
        monitor()
    except findwindows.WindowNotFoundError:
        pass

    while True:
        while not auto_activate():
            pass
        monitor()


# appdata_path = os.environ['APPDATA']
# working_folder = 'SWReminderTool'
# working_dir = f'{appdata_path}/{working_folder}'


# Errors to fix: situation where SW is activated somewhere else
if __name__ == '__main__':
    main()
