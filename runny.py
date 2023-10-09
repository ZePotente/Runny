import os
import os, win32com.client

def get_scripts(path='output'):
    return [os.path.join(dirpath,f) for (dirpath, dirnames, filenames) in os.walk(path) for f in filenames]

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

DEFAULT_EXTENSIONS = ('.exe', '.lnk')
DEFAULT_OUT = 'output'
program_paths = (r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs', r'C:\Users\Pinky\AppData\Roaming\Microsoft\Windows\Start Menu\Programs')
programs_all = [os.path.join(dirpath,f) for p in program_paths for (dirpath, dirnames, filenames) in os.walk(p) for f in filenames]
programs_lnk = [program for program in programs_all if (os.path.splitext(program)[1] in DEFAULT_EXTENSIONS)]

#print('Todos los programas:')
#for p in programs_lnk: print(programs_lnk.index(p), p)

i_selected = [0, 45, 62]
programs_selected = [programs_lnk[i] for i in i_selected]
print('Programas seleccionados:')
print(programs_selected)

out_path = armar_script(programs_selected)
lnk_path = armar_lnk(out_path)

print('Path del script: ' + out_path)
print('Path del link: ' + lnk_path)
print('Listado de scripts:')
print(get_scripts())
