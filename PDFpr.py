from PyPDF2 import PdfFileWriter, PdfFileReader
import PySimpleGUI as Sg
from os import makedirs, listdir, system

main_win_content = [[Sg.Text("Enter Pages (i.e. 0,33,22,69"), Sg.InputText("",
                                                                 key="ranged"),
                     Sg.Button(
        "Select "
                                                              "Folder",
                                                              key="folderButton"), Sg.Button("run", key="run")]]
main_win_layout = [[Sg.Column(layout=main_win_content)]]
main_window = Sg.Window("Super PDF tool", layout=main_win_layout)

while True:
    b, v = main_window.Read(timeout=100)

    final_list = []

    if b == "folderButton":
        folder = Sg.PopupGetFolder("Select folder")
        print("Selected Folder: {}".format(folder))

    if b == "run":
        if not folder:
            Sg.PopupOK("You need to select a folder :)")
            break

        pages = map(int, v['ranged'].split(','))

        for i in pages:
            print("Appending: ", i, " to pages list")
            final_list.append(i)


        for x in listdir(folder):
            hi = PdfFileReader('{}/{}'.format(folder, x), 'rb')
            bye = PdfFileWriter()

            for i in range(0, 69):
                if i in final_list:
                    pass
                else:
                    try:
                        page = hi.getPage(i)
                        bye.addPage(page)
                    except IndexError:
                        pass


            with open('{}/{}'.format(folder, x), 'wb') as f:
                bye.write(f)

        Sg.PopupOK("All Done :D")



    if b == "Exit":
        quit()


