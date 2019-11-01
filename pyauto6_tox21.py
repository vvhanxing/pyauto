from pywinauto import application
import os
import re

app = application.Application()
app.start(r"C:/ACD2018FREE/CHEMSK.EXE")

dlg_spec = app.window(title='ACD/Labs Products')
dlg_spec["OK"].click()

dlg_spec = app.window(title='ACD/ChemSketch (Freeware) - [noname01.sk2]')
dlg_spec.type_keys('%FO')

dlg_spec = app.window(title='Open Document')
dlg_spec.Edit.set_text( "0000.sk2" )
dlg_spec.type_keys('{ENTER}')


    
####################################
def ACD_save_mol(save_mol_name,SMILES_str):
    dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')
    dlg_spec.type_keys('%PD')
    dlg_spec.type_keys('%TGM')

    dlg_spec = app.window(title='Generate Structure from SMILES')
    dlg_spec.Edit.set_text(SMILES_str)
    dlg_spec["OK"].click()


    try:
        dlg_spec = app.window(title=r'Warning')
        dlg_spec.type_keys('Y')
        dlg_spec.type_keys('{ENTER}')
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')
    except:
        
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')


    #dlg_spec.click_input(button='left', coords=(800, 600))

    
    dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')


    dlg_spec.click_input(button='left', coords=(800, 600))

    dlg_spec.type_keys('%T3')

    try:
        dlg_spec = app.window(title=r'3D Structure Optimization')
        dlg_spec.type_keys('Y')
        dlg_spec.type_keys('{ENTER}')
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')
    except:
        
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')



    try:
        dlg_spec = app.window(title=r'Warning')
        dlg_spec.type_keys('Y')
        dlg_spec.type_keys('{ENTER}')
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')
    except:
        
        dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')




        
    dlg_spec.type_keys('%FE')
    dlg_spec = app.window(title=r'ACD/ChemSketch (Freeware) - [C:\Users\Administrator\Desktop\hubChem\lstm_mol\tox21_data\0000.sk2]')
    dlg_spec = app.window(title='Export')


    dlg_spec["保存类型(&T):ComboBox"].select(0)
    dlg_spec.type_keys('%N')
    dlg_spec.Edit.set_text( save_mol_name )
    dlg_spec.type_keys('{ENTER}')
    

####################################


floder = "C:/Users/Administrator/Desktop/hubChem/lstm_mol/tox21/"
atom_type = [ "H","C", "N","O","F","P","S","Cl","Br","I"]



files = os.listdir (floder)


for count ,file in enumerate(files[1934:2100]):
    print(count ,file)
    with  open( floder+ file,"r") as f:
        SMILES_str = f.readlines()[0]
        print(SMILES_str )
    if "." not in SMILES_str and "Cu" not in SMILES_str and "Hg" not in SMILES_str and "Pt" not in SMILES_str and "Bi" not in SMILES_str and  "Ir" not in SMILES_str  and "M" not in SMILES_str:

        ACD_save_mol( file,SMILES_str )
    #if "."  in SMILES_str: 
        #ACD_save_mol( file,SMILES_str )
        
    #os.remove(floder+"0000.mol")
