from tkinter import *


class Host:
    '''Class to keep track of all hosts and perform actions on them.'''
    host_list = []

    def __init__(self, hostname, username, password, mac, column=0, row=0): # ip?
        self.host_list.append(self)
        self.hostname = hostname
        self.username = username
        self.password = password
        self.mac = mac
        self.column = int(column)
        self.row = int(row)
        self.selected = IntVar()
        self.session_id = self.get_session_id()

    def info_frame(self, host_frame):
        '''Return the frame object that gets displayed in the host_frame.'''
        f = Frame(host_frame, padx=5, pady=5, bg="lightgray")
        self.photo = PhotoImage(file="tietokone.png")
        Button(f, image=self.photo, padx=0, pady=0, bg="lightgray", bd=0, command=self.display_properties).pack()
        Checkbutton(f, text=self.hostname, variable=self.selected, bg="lightgray", padx=0, pady=0).pack()
        return f

    def display_properties(self):
        '''Display hosts properties in the properties_frame.'''
        global properties_frame
        for slave in properties_frame.slaves():
            slave.destroy()
        f = Frame(properties_frame, padx=5, pady=5)
        Label(f, text="Hostname: " + self.hostname).pack(anchor=W)
        Label(f, text="Käyttäjänimi: " + self.username).pack(anchor=W)
        #Label(f, text="Salasana: " + self.password).pack(anchor=W)
        Label(f, text="MAC-osoite: " + self.mac).pack(anchor=W)
        Label(f, text="Rivi: " + str(self.row)).pack(anchor=W)
        Label(f, text="Sarake: " + str(self.column)).pack(anchor=W)
        Label(f, text="Sessio ID: " + str(self.session_id)).pack(anchor=W)
        f.pack()

    @classmethod
    def populate(cls, file_path):
        '''Create Host objects from given .csv file.'''
        from csv import reader
        cls.host_list = []
        with open(file_path, "r") as f:
            csv_reader = reader(f, delimiter=" ")
            for row in csv_reader:
                try:
                    Host(*row)
                    print("OK" + str(row))
                except TypeError:
                    print("FAIL: " + str(row))

    @classmethod
    def display(cls):
        '''Display hosts inside the host_frame.

        Hosts with either their row or column defined as 0 get displayed at the bottom on their own row.
        '''
        global host_frame
        for slave in host_frame.grid_slaves():
            slave.destroy()
        undefined_hosts = []
        for host in cls.host_list:
            if host.row == 0 or host.column == 0:
                print("Undefined grid position: " + str(host))
                undefined_hosts.append(host)
            elif host_frame.grid_slaves(row=host.row, column=host.column):
                print("Duplicate grid position: " + str(host))
                undefined_hosts.append(host)
            else:
                host.info_frame(host_frame).grid(row=host.row, column=host.column, padx=5, pady=5)
        number_of_rows = host_frame.grid_size()[1]
        for host in undefined_hosts:
            host.info_frame(host_frame).grid(row=number_of_rows, column=undefined_hosts.index(host)+1, padx=5, pady=5)

    @classmethod
    def run(cls, command, interactive=False, **kwargs):
        '''Run a command on all selected hosts (on a seperate thread to increase performance).'''
        from threading import Thread
        for host in cls.host_list:
            if host.selected.get():
                if interactive:
                    Thread(target=run, args=(host.hostname, host.username, host.password, command, host.session_id), kwargs=kwargs).start()
                else:
                    Thread(target=run, args=(host.hostname, host.username, host.password, command, None), kwargs=kwargs).start()

    @classmethod
    def wake_up(cls):
        '''Wake up all selected hosts through WOL.'''
        from wakeonlan import send_magic_packet
        macs = []
        for host in cls.host_list:
            if host.selected.get(): macs.append(host.mac)
        send_magic_packet(*macs)

    def get_session_id(self):
        '''Get host's active session id.'''
        stdout = run("VKY00093", "uzer", "", "powershell", arguments='-command "Get-Process powershell | Select-Object SessionId"')
        for char in stdout.split():
            if char.isdigit():
                return int(char)

    def get_mac(self):
        '''Get host's mac adress.'''
        from csv import reader
        stdout = run(self.hostname, self.username, self.password, "getmac", arguments="/FO CSV /NH") # Get mac adress from host formatted in csv without the header row
        csv_reader = reader(stdout, delimiter=",")
        return (next(csv_reader)[0]) # Return first entry

    def __repr__(self):
        return "Host('{self.hostname}', '{self.username}', '{self.password}', '{self.mac}', '{self.row}', '{self.column}', '{self.session_id}')".format(self=self)


class Command:
    menus = ["Valitse", "Tietokone", "VBS3", "SteelBeasts", "Muut"]

    def __init__(self, name):
        self.name = name

    def clicked(self):
        '''Method which should be called by the children of the Command class to get the basic functionality of setting up the properties frame.'''
        global properties_frame
        for slave in properties_frame.slaves():
            slave.destroy()
        f = Frame(properties_frame, padx=5, pady=5)
        f.pack()
        return f

    @classmethod
    def init_commands(cls, menubar):
        '''Creates menu object from the menu labels specified in the "menus" class variable.'''
        for menu_label in cls.menus:
            menu = Menu(menubar, tearoff=0)
            cls.menus[cls.menus.index(menu_label)] = menu
            menubar.add_cascade(label=menu_label, menu=menu)

    def add_to_menu(self, menu_index):
        self.menus[menu_index].add_command(label=self.name, command=self.clicked)


class BatchCommand(Command):
    def __init__(self, name, command, interactive=False, **kwargs):
        super().__init__(name)
        self.command = command
        self.interactive = interactive
        self.kwargs = kwargs

    def clicked(self):
        Host.run(self.command, self.interactive, **self.kwargs)


class SelectAll(Command):
    def clicked(self):
        for host in Host.host_list:
            host.selected.set(1)


class Deselect(Command):
    def clicked(self):
        for host in Host.host_list:
            host.selected.set(0)


class InvertSelection(Command):
    def clicked(self):
        for host in Host.host_list:
            if host.selected.get():
                host.selected.set(0)
            else:
                host.selected.set(1)


class SelectX(Command):
    def __init__(self, name):
        super().__init__(name)
        self.column = StringVar()
        self.row = StringVar()

    def clicked(self):
        f = super().clicked()
        Label(f, text="Sarake:").pack(anchor=W)
        Entry(f, textvariable=self.column, width=10).pack(anchor=W)
        Button(f, text="Valitse", command=self.select_column).pack(anchor=W)
        Label(f, text="Rivi:").pack(anchor=W)
        Entry(f, textvariable=self.row, width=10).pack(anchor=W)
        Button(f, text="Valitse", command=self.select_row).pack(anchor=W)

    def select_column(self):
        for host in Host.host_list:
            if host.column == int(self.column.get()):
                host.selected.set(1)

    def select_row(self):
        for host in Host.host_list:
            if host.row == int(self.row.get()):
                host.selected.set(1)


class CustomCommand(Command):
    def __init__(self, name, command="", arguments="", interactive=0, can_be_changed=True, **kwargs):
        super().__init__(name)
        self.command = StringVar()
        self.command.set(command)
        self.arguments = StringVar()
        self.arguments.set(arguments)
        self.interactive = IntVar()
        self.interactive.set(interactive)
        self.can_be_changed = can_be_changed
        self.kwargs = kwargs

    def clicked(self):
        f = super().clicked()
        Label(f, text="Komento:").pack(anchor=W)
        Entry(f, textvariable=self.command, width=50).pack(anchor=W)
        Label(f, text="Parametrit:").pack(anchor=W)
        Entry(f, textvariable=self.arguments, width=50).pack(anchor=W)
        if self.can_be_changed:
            Checkbutton(f, text="Interaktiivinen", variable=self.interactive).pack()
        Button(f, text="Aja", command=self.run).pack(anchor=W)

    def run(self):
        Host.run(self.command.get(), self.interactive.get(), arguments=self.arguments.get(), **self.kwargs)


class UpdateCommand(Command):
    def __init__(self, name, default_file_path):
        super().__init__(name)
        self.file_path = StringVar()
        self.file_path.set(default_file_path)

    def clicked(self):
        f = super().clicked()
        Label(f, text="Luokkatiedoston polku:").pack(anchor=W)
        Entry(f, textvariable=self.file_path, width=50).pack(anchor=W)
        Button(f, text="Päivitä", command=self.run).pack(anchor=W)
    
    def run(self):
        Host.populate(self.file_path.get())
        Host.display()


class BootCommand(Command):
    def clicked(self):
        Host.wake_up()


class CopyCommand(Command):
    def __init__(self, name, source, destination, **kwargs):
        super().__init__(name)
        self.source = StringVar()
        self.source.set(source)
        self.destination = StringVar()
        self.destination.set(destination)
        self.kwargs = kwargs

    def clicked(self):
        f = super().clicked()
        Label(f, text="Lähde:").pack(anchor=W)
        Entry(f, textvariable=self.source, width=50).pack(anchor=W)
        Label(f, text="Kohde:").pack(anchor=W)
        Entry(f, textvariable=self.destination, width=50).pack(anchor=W)
        Button(f, text="Kopioi", command=self.run).pack(anchor=W)

    def run(self):
        args = self.source.get() + " " + self.destination.get()
        Host.run("robocopy", interactive=False, arguments=args, **self.kwargs)


def run(host, username, password, command, session_id=None, print_std=True, **kwargs):
    '''Run a command on a specific host.'''
    from pypsexec.client import Client
    from pypsexec.exceptions import PAExecException
    from smbprotocol.exceptions import SMBAuthenticationError
    from socket import gaierror
    c = Client(host, username=username, password=password, encrypt=False)
    try:
        c.connect()
    except SMBAuthenticationError as e:
        print("Autentikointi virhe - " + str(e))
        return
    except gaierror as e:
        print("Virheellinen kohde - " + str(e))
        return
    try:
        c.create_service()
        if session_id:
            stdout, stderr, pid = c.run_executable(command, interactive=True, interactive_session=session_id, use_system_account=True, asynchronous=True, **kwargs)
            print("'{}' started on {} with PID {}".format(command, host, pid))
        else:
            stdout, stderr, pid = c.run_executable(command, **kwargs)
            if print_std:
                print(stderr.decode(encoding='windows-1252'))
                print(stdout.decode(encoding='windows-1252'))
            return stdout.decode(encoding='windows-1252')
    except PAExecException as e:
        print("Virheellinen komento - " + str(e))
    finally:
        c.remove_service()
        c.disconnect()


def main():
    '''Entry point of the program.'''
    from os.path import normpath
    global host_frame, properties_frame
    root = Tk()
    root.title("Luokanhallinta")
    root.geometry("1280x720")
    host_frame = LabelFrame(root, text="Luokka")
    host_frame.pack(side=LEFT, fill=BOTH, expand="yes")
    properties_frame = LabelFrame(root, text="Ominaisuudet")
    properties_frame.pack(side=RIGHT, fill=BOTH)
    menubar = Menu(root)
    Command.init_commands(menubar)
    root.config(menu=menubar)

    SelectAll("Kaikki").add_to_menu(0)
    Deselect("Ei mitään").add_to_menu(0)
    InvertSelection("Käänteinen").add_to_menu(0)
    SelectX("Valitse tietty...").add_to_menu(0)

    BootCommand("Käynnistä").add_to_menu(1)
    BatchCommand("Käynnistä uudelleen", "shutdown", arguments="/r").add_to_menu(1)
    BatchCommand("Sammuta", "shutdown", arguments="/s").add_to_menu(1)

    CustomCommand("Käynnistä...", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\VBS3_64.exe", interactive=1, can_be_changed=False).add_to_menu(2)
    CopyCommand("Synkkaa Addonit...", normpath("//PSPR-Storage/Addons"), normpath('"C:/Program Files/Bohemia Interactive Simulations/VBS3 3.9.0.FDF EZYQC_FI/mycontent/addons"')).add_to_menu(2)
    BatchCommand("Sulje", "taskkill", arguments="/im VBS3_64.exe /F").add_to_menu(2)

    CustomCommand("Käynnistä...", "C:\Program Files\eSim Games\SB Pro FI\Release\SBPro64CM.exe", interactive=1, can_be_changed=False).add_to_menu(3)
    BatchCommand("Sulje", "taskkill", arguments="/im SBPro64CM.exe /F").add_to_menu(3)

    UpdateCommand("Päivitä luokka...", "luokka.csv").add_to_menu(4)
    CopyCommand("Siirrä tiedostoja...", "", "").add_to_menu(4)
    CustomCommand("Aja...").add_to_menu(4)

    Host.populate(normpath("luokka.csv")) 
    Host.display()
    root.mainloop()


if __name__ == "__main__":
    main()