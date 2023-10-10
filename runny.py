import os
import win32com.client

def get_scripts(dir_path='output'):
    """Devuelve una lista de los paths de los scripts en la carpeta path"""
    return [os.path.join(dirpath,f) for (dirpath, dirnames, filenames) in os.walk(dir_path) for f in filenames]

def armar_script(programs, nombre_script='script.bat'):
    """Arma un script a partir de la lista de programas y devuelve el path del script"""
    
    output_script = os.path.join(DEFAULT_OUT, nombre_script)
    
    with open(output_script, 'w') as f:
        f.write(r'@echo off' + '\n')
        for program in programs:
            f.write('start "" "' + program + '"\n')
        f.write('exit')
    
    return output_script

def armar_lnk(path_script, nombre_lnk='script.lnk', args='', hotkey = ''):
    """Arma un archivo .lnk en el escritorio y devuelve el path del .lnk"""
    
    path_lnk = os.path.join(r'~\Desktop', nombre_lnk)
    
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(os.path.expanduser(path_lnk))
    shortcut.Targetpath = os.path.join(os.getcwd(), path_script)
    shortcut.WorkingDirectory = os.getcwd()
    shortcut.Arguments = args
    shortcut.Hotkey = hotkey
    #shortcut.IconLocation = icon
    shortcut.save()

    return path_lnk

def obtener_path_programas(prog, nombres):
    return [p for p in prog for nombre in nombres if nombre == nombre_from_path(p)]

def nombre_from_path(path):
    return os.path.splitext(os.path.split(path)[1])[0]

DEFAULT_EXTENSIONS = ('.exe', '.lnk')
DEFAULT_OUT = 'output'
program_paths = (r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs', 
                 r'C:\Users\Pinky\AppData\Roaming\Microsoft\Windows\Start Menu\Programs',
                 #os.path.expanduser(r'~\Desktop') # Rompe la unicidad de los programas
                 )
programs_all = [os.path.join(dirpath,f) for p in program_paths for (dirpath, dirnames, filenames) in os.walk(p) for f in filenames]
programs_lnk = [program for program in programs_all if (os.path.splitext(program)[1] in DEFAULT_EXTENSIONS)]

#print('Todos los programas:')
#for p in programs_lnk: print(programs_lnk.index(p), p)

i_selected = [0, 45, 62]
programs_selected = [programs_lnk[i] for i in i_selected]
#print('Programas seleccionados:')
#print(programs_selected)

out_path = armar_script(programs_selected)
lnk_path = armar_lnk(out_path)

#print('Path del script: ' + out_path)
#print('Path del link: ' + lnk_path)
#print('Listado de scripts:')
#print(get_scripts())

# No digan cómo programo por favor
## GUI
import PySimpleGUI as sg
nombres_selected = []
#programas = sorted([os.path.splitext(os.path.split(p)[1])[0] + ' - ' + p for p in sorted(programs_lnk)])
programas = sorted([os.path.splitext(os.path.split(p)[1])[0] for p in sorted(programs_lnk)])
program_list_column = [
    # Buscador
    [sg.Text('Programas')],
    [
        sg.Text('Busqueda:'), 
        sg.Input(enable_events=True, key='-IN-', size=(30,None))
    ],
    [sg.Listbox(
        values=programas,
        enable_events=False, size=(40,20), horizontal_scroll=True,
        key="-LISTA PROGRAMAS-"
    )],
    [sg.Button('Agregar')]
]

shortcut_list_column = [
    [sg.Text('Shortcut')],
    [sg.Listbox(
        values=nombres_selected, enable_events=True, size=(40,20), horizontal_scroll=True,
        key='-LISTA SHORTCUT-'
    )],
    [
        sg.Button('Quitar'), 
        sg.Button('Limpiar'), 
        #sg.Button('Editar')
    ],
    [
        sg.Text('Nombre:'), 
        sg.Input(enable_events=True, key='-OUT-', size=(25,None)),
        sg.Button('Crear')]
]

layout = [
    [
        sg.Column(program_list_column),
        sg.VSeparator(),
        sg.Column(shortcut_list_column)
    ]
]

window = sg.Window('Runny-GUI', layout)

# event loop
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    # Pestaña Programas
    if event == 'Agregar':
        nombres_selected.extend(window['-LISTA PROGRAMAS-'].get())
        window['-LISTA SHORTCUT-'].update(nombres_selected)
    if event == '-IN-':
        busqueda = values['-IN-'].casefold()
        window['-LISTA PROGRAMAS-'].update([p for p in programas if busqueda in p.casefold()])
    # Pestaña Shortcut
    if event == 'Quitar':
        if (window['-LISTA SHORTCUT-'].get() != []):
            p = window['-LISTA SHORTCUT-'].get()[0]
            nombres_selected.remove(p)
            window['-LISTA SHORTCUT-'].update(nombres_selected)
    if event == 'Crear':
        if nombres_selected != []:
            #nombre_lnk = 'hola.lnk'
            nombre_lnk = os.path.splitext(values['-OUT-'])[0]
            programas_selected = obtener_path_programas(programs_lnk, nombres_selected)
            out_path = armar_script(programas_selected, nombre_script=nombre_lnk + '.bat')
            lnk_path = armar_lnk(out_path, nombre_lnk=nombre_lnk + '.lnk')
            # Tirar popup de éxito
            sg.popup('Atajo creado exitosamente')
    if event == 'Limpiar':
        nombres_selected.clear()
        window['-LISTA SHORTCUT-'].update(nombres_selected)

window.close()
