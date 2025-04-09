import win32com.client

def creer_et_executer_macro(fichier_excel):
    # Ouvrir l'application Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Optionnel : rendre Excel visible pour le débogage
    
    # Ouvrir le fichier Excel
    workbook = excel.Workbooks.Open(fichier_excel)
    
    # Ajouter un module VBA et insérer une macro
    module = workbook.VBProject.VBComponents.Add(1)  # 1 = Module standard
    macro = """
    Sub TestMacro()
        MsgBox "Bonjour depuis VBA !"
    End Sub
    """
    module.CodeModule.AddFromString(macro)
    
    # Exécuter la macro
    excel.Application.Run("TestMacro")
    
    # Sauvegarder et fermer le fichier Excel
    workbook.Save()
    workbook.Close()
    excel.Quit()
    print("Macro exécutée et fichier sauvegardé.")

# Exemple d'utilisation
fichier_excel = r"chemin\vers\fichier.xlsx"  # Remplacez par le chemin réel de votre fichier Excel
creer_et_executer_macro(fichier_excel)
