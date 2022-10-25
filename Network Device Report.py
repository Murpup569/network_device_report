#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
REPORT
This outputs a csv file with the hostname of the switch and today's date.
This script uses netmiko to gather information via ssh.
---Order of commands---
'show ip interface brief | ex OK'
'show interfaces description | exclude Protocol Description'
'show cdp nei' + ip_int_br[i][0]'
'show int {ip_int_br[i][0]} | in Last input'
'show int ' + ip_int_br[i][0] + ' capabilities | in Type|Duplex'
'show int ' + ip_int_br[i][0] + ' switchport | in Administrative Mode|Operational Mode|Access Mode VLAN'
'show mac address-table interface ' + ip_int_br[i][0] + ' | ex Vlan|-|Table|Total'
'show run | in hostname'

TODO show int | in error|drops|GigabitEthernet
"""

__author__ = 'Ryan Murray'
__version__ = '2.0'
__maintainer__ = 'Ryan Murray'
__email__ = 'ryan.murray.570@gmail.com'
__contributors__ = 'Ryan Murray'

import re
import sys
import traceback
from datetime import date
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from tkinter import ttk
from socket import gethostbyaddr

import openpyxl
import pandas as pd
from mac_vendor_lookup import MacLookup
from netmiko import ConnectHandler
from styleframe import StyleFrame

def show_error(self, *args):
    err = traceback.format_exception(*args)
    messagebox.showerror('Exception', err)

    L2DeviceFrame = ttk.LabelFrame(root, text='Layer 2 Device', padding=(10,10,10,10))
    L2DeviceFrame.grid(column=0, row=0, padx=10, pady=10, columnspan=2)
    L3DeviceFrame = ttk.LabelFrame(root, text='Layer 3 Device', padding=(10,10,10,10))
    L3DeviceFrame.grid(column=2, row=0, padx=10, pady=10, columnspan=2)

    L2IPLabel = Label(L2DeviceFrame, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    L2usernameLabel = Label(L2DeviceFrame, text='Enter Username: ', padx=10, width=25, anchor=W)
    L2passwordLabel = Label(L2DeviceFrame, text='Enter Password: ', padx=10, width=25, anchor=W)

    L2ipAddress = Entry(L2DeviceFrame)
    L2username = Entry(L2DeviceFrame)
    L2password = Entry(L2DeviceFrame, show='*')
    L2username.insert(0, "admin")

    L3IPLabel = Label(L3DeviceFrame, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    L3usernameLabel = Label(L3DeviceFrame, text='Enter Username: ', padx=10, width=25, anchor=W)
    L3passwordLabel = Label(L3DeviceFrame, text='Enter Password: ', padx=10, width=25, anchor=W)

    L3ipAddress = Entry(L3DeviceFrame)
    L3username = Entry(L3DeviceFrame)
    L3password = Entry(L3DeviceFrame, show='*')
    L3ipAddress.insert(0, "10.254.100.2")
    L3username.insert(0, "admin")

    L3Enable = Button(L3DeviceFrame, text='Enable', command=enable_l3)
    L3Disable = Button(L3DeviceFrame, text='Disable', command=disable_l3)

    no_po_selected = BooleanVar()
    no_po_selected.set(True)
    po_channel = Checkbutton(L2DeviceFrame, text="Do not gather port-channel macs", variable=no_po_selected)
    show = Label(root, text='                       ')
    submitButton = Button(root, text='Submit', command=lambda: submit(L2ipAddress.get(), L2username.get(), L2password.get(), L3ipAddress.get(), L3username.get(), L3password.get(), no_po_selected.get()))
    closeButton = Button(root, text='Close', command=lambda: close(root))

    L2IPLabel.grid(row=0, column=0)
    L2usernameLabel.grid(row=1, column=0)
    L2passwordLabel.grid(row=2, column=0)

    L2ipAddress.grid(row=0, column=1)
    L2username.grid(row=1, column=1)
    L2password.grid(row=2, column=1)
    po_channel.grid(row=3, column=0, columnspan=2)

    L3IPLabel.grid(row=0, column=2)
    L3usernameLabel.grid(row=1, column=2)
    L3passwordLabel.grid(row=2, column=2)
    L3ipAddress.grid(row=0, column=3)
    L3username.grid(row=1, column=3)
    L3password.grid(row=2, column=3)
    L3Enable.grid(row=3, column=2)
    L3Disable.grid(row=3, column=3)

    submitButton.grid(row=4, column=0)
    closeButton.grid(row=4, column=1)
    show.grid(row=5, column=0)

    root.update()

def close(root):
    result = messagebox.askquestion(title="Closeing", message="Are you sure you want to close?", icon="warning")
    if result == "yes":
        Tk.report_callback_exception = print()
        try:
            # Disconnects from device
            l2_net_connect.disconnect()
            if GatherL3Info:
                l3_net_connect.disconnect()
        except:
            pass
        root.destroy()
        sys.exit()

def submit(L2ipAddress, L2username, L2password, L3ipAddress, L3username, L3password, no_po_selected):
    global show
    global progress
    global l2_net_connect
    global l2_net_device
    global l3_net_connect
    global l3_net_device
    global GatherL3Info

    show.destroy()
    show = Label(root, text=f'Atempting to connect to {L2ipAddress}', padx=10)
    show.grid(row=5, column=0)
    root.update()

    l2_net_device = {
        'device_type': 'cisco_ios',
        'ip': L2ipAddress,
        'username': L2username,
        'password': L2password,
    }

    if GatherL3Info:
        l3_net_device = {
            'device_type': 'cisco_ios',
            'ip': L3ipAddress,
            'username': L3username,
            'password': L3password,
        }

    # Logs into the networking device
    l2_net_connect = ConnectHandler(**l2_net_device)
    if GatherL3Info:
        l3_net_connect = ConnectHandler(**l3_net_device)
    show.destroy()
    show = Label(root, text=f'Connection to {L2ipAddress} successful!', padx=10)
    show.grid(row=5, column=0)
    root.update()

    # Enters 'show ip int br' and puts it in ip_int_br
    show_ip_int_br = l2_net_connect.send_command('show ip interface brief | ex OK')
    show_ip_int_br = show_ip_int_br.lstrip('\n')
    show_ip_int_br = show_ip_int_br.rstrip('\n')
    ip_int_br = [x.split() for x in show_ip_int_br.split('\n')]

    # Enters 'show int desc' and puts it in int_desc
    show_int_desc = l2_net_connect.send_command('show interfaces description | exclude Protocol Description')
    int_desc = []
    for x in show_int_desc.split('\n'):
        int_desc.append(x[55:])

    # Defines the table to use and inputs headers
    table = []

    for i in range(len(ip_int_br)):
        show.destroy()
        show = Label(root, text=f'Processing port {str(i)} out of {str(len(ip_int_br))}', padx=10)
        show.grid(row=5, column=0)
        MAX = int(len(ip_int_br))
        progress_var = DoubleVar()
        progress = Progressbar(root, orient=HORIZONTAL, length=200, variable=progress_var, maximum=MAX)
        progress_var.set(i)
        progress.grid(row=6, column=0, padx=10, pady=(0, 10))
        root.update()

        int_status = ip_int_br[i][4] + '/' + ip_int_br[i][5]

        # If the interface is a vlan it will not try to find cdp nei, speed, duplex, switchport info, mac, and oui lookup
        if re.search('Vlan.+', ip_int_br[i][0]):

            # Adds gathered information to table without mac
            if GatherL3Info:
                table.append([ip_int_br[i][0],                   # Interface
                    '',                                          # Speed
                    '',                                          # Duplex
                    '',                                          # Switchport
                    '',                                          # Vlan
                    ip_int_br[i][1],                             # SVI IP Address
                    int_status,                                  # Status
                    '',                                          # Connected Mac
                    '',                                          # Client IP Address
                    '',                                          # Client Hostname
                    '',                                          # OUI Lookup
                    int_desc[i],                                 # Description
                    ''                                           # CDP Neighbors
                    ''                                           # Last Input
                ])
            else:
                table.append([ip_int_br[i][0],                   # Interface
                    '',                                          # Speed
                    '',                                          # Duplex
                    '',                                          # Switchport
                    '',                                          # Vlan
                    ip_int_br[i][1],                             # SVI IP Address
                    int_status,                                  # Status
                    '',                                          # Connected Mac
                    '',                                          # OUI Lookup
                    int_desc[i],                                 # Description
                    ''                                           # CDP Neighbors
                    ''                                           # Last Input
                ])
        else:
            
            # Gathers CDP information
            if int_status != 'down/down':
                cdp_nei = l2_net_connect.send_command(f'show cdp nei {ip_int_br[i][0]}')
                cdp_nei = cdp_nei.lstrip('\n')
                cdp_nei = cdp_nei[289:]
                cdp_nei = cdp_nei.split(' ')
                cdp_nei = cdp_nei[0].strip()
                cdp_nei = ('' if cdp_nei == 'Total' else cdp_nei)
            else:
                cdp_nei = ''
            
            # Gathers Last Input
            if int_status == 'up/up':
                last_input = ''
            else:
                last_input = l2_net_connect.send_command(f'show int {ip_int_br[i][0]} | in Last input')
                p0 = re.compile(r'Last input (?P<time>.+), output [^h]')
                last_input = last_input.strip()
                m = p0.match(last_input)
                last_input = m.groupdict()['time']
                last_input = ('' if last_input == 'never' else last_input)

            # Gathers Speed and Duplex information 
            show_speed_duplex = l2_net_connect.send_command(f'show int {ip_int_br[i][0]} capabilities | in Type|Duplex')
            show_speed_duplex = show_speed_duplex[1:]
            speed_duplex = [x.split() for x in show_speed_duplex.split('\n')]
            try:
                speed = speed_duplex[0][1]
                duplex = speed_duplex[1][1]
            except:
                speed = ''
                duplex = ''

            # Gathers Etherchannel and Trunk information
            show_switchport = l2_net_connect.send_command(f'show int {ip_int_br[i][0]} switchport | in Administrative Mode|Operational Mode|Access Mode VLAN')
            show_switchport = show_switchport.lstrip('\n')
            show_switchport = show_switchport.rstrip('\n')
            switchport = show_switchport.split('\n')
            trunk_access = switchport[0]
            trunk_access = trunk_access[21:]
            if trunk_access == 'trunk' or trunk_access == ' trunk':
                trunk_access = switchport[1]
                trunk_access = trunk_access[18:]
            try:
                vlan = switchport[2]
                vlan = vlan[18:]
            except:
                vlan = 'not switchable'

            # If the user checkmarked do not gather port-channel macs or if the interface is down
            if (no_po_selected and re.search('Port-channel.+', ip_int_br[i][0])) or int_status == 'down/down':
                if GatherL3Info:
                    table.append([ip_int_br[i][0],                   # Interface
                        speed,                                       # Speed
                        duplex,                                      # Duplex
                        trunk_access,                                # Switchport
                        vlan,                                        # Vlan
                        ip_int_br[i][1],                             # IP Address
                        int_status,                                  # Status
                        '',                                          # Connected Mac
                        '',                                          # Client IP Address
                        '',                                          # Client Hostname
                        '',                                          # OUI Lookup
                        int_desc[i],                                 # Description
                        cdp_nei,                                     # CDP Neighbors
                        last_input                                   # Last Input
                    ])
                else:
                    table.append([ip_int_br[i][0],                   # Interface
                        speed,                                       # Speed
                        duplex,                                      # Duplex
                        trunk_access,                                # Switchport
                        vlan,                                        # Vlan
                        ip_int_br[i][1],                             # IP Address
                        int_status,                                  # Status
                        '',                                          # Connected Mac
                        '',                                          # OUI Lookup
                        int_desc[i],                                 # Description
                        cdp_nei,                                     # CDP Neighbors
                        last_input                                   # Last Input
                    ])
            else:
                # Gathers mac address information
                mac_table = l2_net_connect.send_command(f'show mac address-table interface {ip_int_br[i][0]} | ex Vlan|-|Table|Total')
                mac_table = mac_table.lstrip('\n')
                mac_table = mac_table.rstrip('\n')
                mac_table = mac_table.split('\n')
                for m in range(len(mac_table)):
                    root.update()
                    mac = mac_table[m].split()
                    try:
                        mac = mac[1]
                        oui = MacLookup().lookup(str(mac))
                        if GatherL3Info:
                            try:
                                arpTable = l3_net_connect.send_command(f'show ip arp | in {mac}')
                                arpTable = arpTable.strip("\n")
                                if arpTable:
                                    p1 = re.compile(r'Internet  (?P<client_ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}).+')
                                    if "\n" in arpTable:
                                        arpTable = arpTable.split("\n")
                                        for line in arpTable:
                                            m = p1.match(line)
                                            if m:
                                                client_ip = m.groupdict()['client_ip']
                                                try:
                                                    client_hostname = gethostbyaddr(client_ip)[0]
                                                except:
                                                    client_hostname = ''
                                                table.append([ip_int_br[i][0],                   # Interface
                                                    speed,                                       # Speed
                                                    duplex,                                      # Duplex
                                                    trunk_access,                                # Switchport
                                                    vlan,                                        # Vlan
                                                    ip_int_br[i][1],                             # SVI IP Address
                                                    int_status,                                  # Status
                                                    mac,                                         # Connected Mac
                                                    client_ip,                                   # Client IP Address
                                                    client_hostname,                             # Client Hostname
                                                    oui,                                         # OUI Lookup
                                                    int_desc[i],                                 # Description
                                                    cdp_nei,                                     # CDP Neighbors
                                                    last_input                                   # Last Input
                                                ])
                                    else:
                                        m = p1.match(arpTable)
                                        if m:
                                            client_ip = m.groupdict()['client_ip']
                                            try:
                                                client_hostname = gethostbyaddr(client_ip)[0]
                                            except:
                                                client_hostname = ''
                                            table.append([ip_int_br[i][0],                   # Interface
                                                speed,                                       # Speed
                                                duplex,                                      # Duplex
                                                trunk_access,                                # Switchport
                                                vlan,                                        # Vlan
                                                ip_int_br[i][1],                             # SVI IP Address
                                                int_status,                                  # Status
                                                mac,                                         # Connected Mac
                                                client_ip,                                   # Client IP Address
                                                client_hostname,                             # Client Hostname
                                                oui,                                         # OUI Lookup
                                                int_desc[i],                                 # Description
                                                cdp_nei,                                     # CDP Neighbors
                                                last_input                                   # Last Input
                                            ])
                            except:
                                table.append([ip_int_br[i][0],                   # Interface
                                    speed,                                       # Speed
                                    duplex,                                      # Duplex
                                    trunk_access,                                # Switchport
                                    vlan,                                        # Vlan
                                    ip_int_br[i][1],                             # SVI IP Address
                                    int_status,                                  # Status
                                    mac,                                         # Connected Mac
                                    '',                                          # Client IP Address
                                    '',                                          # Client Hostname
                                    oui,                                         # OUI Lookup
                                    int_desc[i],                                 # Description
                                    cdp_nei,                                     # CDP Neighbors
                                    last_input                                   # Last Input
                                ])
                        else:
                            table.append([ip_int_br[i][0],                   # Interface
                                speed,                                       # Speed
                                duplex,                                      # Duplex
                                trunk_access,                                # Switchport
                                vlan,                                        # Vlan
                                ip_int_br[i][1],                             # SVI IP Address
                                int_status,                                  # Status
                                mac,                                         # Connected Mac
                                oui,                                         # OUI Lookup
                                int_desc[i],                                 # Description
                                cdp_nei,                                     # CDP Neighbors
                                last_input                                   # Last Input
                            ])
                    except:
                        mac = ''
                        oui = ''
                        if GatherL3Info:
                            l3_net_connect.send_command(f' ')
                            table.append([ip_int_br[i][0],                   # Interface
                                speed,                                       # Speed
                                duplex,                                      # Duplex
                                trunk_access,                                # Switchport
                                vlan,                                        # Vlan
                                ip_int_br[i][1],                             # SVI IP Address
                                int_status,                                  # Status
                                mac,                                         # Connected Mac
                                '',                                          # Client IP Address
                                '',                                          # Client Hostname
                                oui,                                         # OUI Lookup
                                int_desc[i],                                 # Description
                                cdp_nei,                                     # CDP Neighbors
                                last_input                                   # Last Input
                            ])
                        else:
                            table.append([ip_int_br[i][0],                   # Interface
                                speed,                                       # Speed
                                duplex,                                      # Duplex
                                trunk_access,                                # Switchport
                                vlan,                                        # Vlan
                                ip_int_br[i][1],                             # SVI IP Address
                                int_status,                                  # Status
                                mac,                                         # Connected Mac
                                oui,                                         # OUI Lookup
                                int_desc[i],                                 # Description
                                cdp_nei,                                     # CDP Neighbors
                                last_input                                   # Last Input
                            ])

    # Gathers hostname from network device
    hostname = l2_net_connect.send_command('show run | in hostname')
    hostname = hostname[9:]

    # Disconnects from device
    l2_net_connect.disconnect()
    if GatherL3Info:
        l3_net_connect.disconnect()

    show.destroy()
    show = Label(root, text=f'Saving report as {hostname} {str(date.today())}.xlsx')
    show.grid(row=5, column=0)
    root.update()

    # Save table as a xlsx
    StyleFrame.A_FACTOR = 3
    if GatherL3Info:
        columns = ['Interface','Speed','Duplex','Switchport','Vlan','SVI IP Address','Status','Connected Mac','IP Address','Hostname','OUI Lookup','Description','CDP Neighbors','Last Input']
    else:
        columns = ['Interface','Speed','Duplex','Switchport','Vlan','SVI IP Address','Status','Connected Mac','OUI Lookup','Description','CDP Neighbors','Last Input']
    export_file_path = filedialog.asksaveasfilename(initialfile=hostname + ' ' + str(date.today()),defaultextension='.xlsx')
    df = pd.DataFrame(data=table, columns=columns)
    excel_writer = StyleFrame.ExcelWriter(export_file_path)
    sf = StyleFrame(df)
    sf.to_excel(
        excel_writer=excel_writer, 
        best_fit=columns,
        columns_and_rows_to_freeze='B2', 
        row_to_add_filters=0,
    )
    excel_writer.save()

    messagebox.showinfo(title='Finished!', message='Time to party!')

def enable():
    global GatherL3Info
    GatherL3Info = True

    L3IPLabel.config(state='active')
    L3usernameLabel.config(state='active')
    L3passwordLabel.config(state='active')

    L3ipAddress.config(state='normal')
    L3username.config(state='normal')
    L3password.config(state='normal')

    L3Disable.config(state='active')
    L3Enable.config(state='disabled')

def disable():
    global GatherL3Info
    GatherL3Info = False

    L3IPLabel.config(state='disabled')
    L3usernameLabel.config(state='disabled')
    L3passwordLabel.config(state='disabled')

    L3ipAddress.config(state='disabled')
    L3username.config(state='disabled')
    L3password.config(state='disabled')

    L3Disable.config(state='disabled')
    L3Enable.config(state='active')

if __name__ == '__main__':
    root = Tk()
    root.title('Network Device Report')

    Tk.report_callback_exception = show_error

    global GatherL3Info
    GatherL3Info = False

    L2DeviceFrame = ttk.LabelFrame(root, text='Layer 2 Device', padding=(10,10,10,10))
    L2DeviceFrame.grid(column=0, row=0, padx=10, pady=10, columnspan=2)
    L3DeviceFrame = ttk.LabelFrame(root, text='Layer 3 Device', padding=(10,10,10,10))
    L3DeviceFrame.grid(column=2, row=0, padx=10, pady=10, columnspan=2)

    L2IPLabel = Label(L2DeviceFrame, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    L2usernameLabel = Label(L2DeviceFrame, text='Enter Username: ', padx=10, width=25, anchor=W)
    L2passwordLabel = Label(L2DeviceFrame, text='Enter Password: ', padx=10, width=25, anchor=W)

    L2ipAddress = Entry(L2DeviceFrame)
    L2username = Entry(L2DeviceFrame)
    L2password = Entry(L2DeviceFrame, show='*')
    L2username.insert(0, "admin")

    L3IPLabel = Label(L3DeviceFrame, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    L3usernameLabel = Label(L3DeviceFrame, text='Enter Username: ', padx=10, width=25, anchor=W)
    L3passwordLabel = Label(L3DeviceFrame, text='Enter Password: ', padx=10, width=25, anchor=W)

    L3ipAddress = Entry(L3DeviceFrame)
    L3username = Entry(L3DeviceFrame)
    L3password = Entry(L3DeviceFrame, show='*')
    L3username.insert(0, "admin")

    L3Enable = Button(L3DeviceFrame, text='Enable', command=enable)
    L3Disable = Button(L3DeviceFrame, text='Disable', command=disable)

    no_po_selected = BooleanVar()
    no_po_selected.set(True)
    po_channel = Checkbutton(L2DeviceFrame, text="Do not gather port-channel macs", variable=no_po_selected)
    show = Label(root, text='                       ')
    submitButton = Button(root, text='Submit', command=lambda: submit(L2ipAddress.get(), L2username.get(), L2password.get(), L3ipAddress.get(), L3username.get(), L3password.get(), no_po_selected.get()))
    closeButton = Button(root, text='Close', command=lambda: close(root))

    L2IPLabel.grid(row=0, column=0)
    L2usernameLabel.grid(row=1, column=0)
    L2passwordLabel.grid(row=2, column=0)

    L2ipAddress.grid(row=0, column=1)
    L2username.grid(row=1, column=1)
    L2password.grid(row=2, column=1)
    po_channel.grid(row=3, column=0, columnspan=2)

    L3IPLabel.grid(row=0, column=2)
    L3usernameLabel.grid(row=1, column=2)
    L3passwordLabel.grid(row=2, column=2)
    L3ipAddress.grid(row=0, column=3)
    L3username.grid(row=1, column=3)
    L3password.grid(row=2, column=3)
    L3Enable.grid(row=3, column=2)
    L3Disable.grid(row=3, column=3)

    L3IPLabel.config(state='disabled')
    L3usernameLabel.config(state='disabled')
    L3passwordLabel.config(state='disabled')

    L3ipAddress.config(state='disabled')
    L3username.config(state='disabled')
    L3password.config(state='disabled')

    L3Disable.config(state='disabled')

    submitButton.grid(row=4, column=0)
    closeButton.grid(row=4, column=1)
    show.grid(row=5, column=0)

    root.mainloop()
