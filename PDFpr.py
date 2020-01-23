import PyPDF2
from PyPDF2.utils import PyPdfError, PdfReadWarning, PdfReadError, \
    PageSizeNotDefinedError, PdfStreamError
import PySimpleGUI as Sg
from os import makedirs, listdir, system


folder = ""
main_win_content = [[Sg.Text("Pages to Remove:"),
                     Sg.InputText("", key="ranged")],
                     [Sg.Text("Folder:                "), Sg.InputText("",
                                                            key="folderloc",
                                                       disabled=True),
                      Sg.Button("Search",
                                               key="folderButton")],
                     [Sg.Text("")],
                     [Sg.Text("* Pages delimiter is comma with no spaces"
                              " i.e. 1,3,69 ")],
                     [Sg.Text("* Clicking Run will remove pages from all "
                              "PDF's\n   from within the selected folder "
                              "directory.\n   Make backups prior to use.")],
                     [Sg.Text("")],
                     [Sg.Button("Run", key="run")],[Sg.Text("")],
                        [Sg.ProgressBar(69, orientation='h', size=(50,
                                                                     200),
                                        key='progressbar')]]

main_win_layout = [[Sg.Column(layout=main_win_content)]]

main_window = Sg.Window("Batch PDF Page Remover",
                        layout=main_win_layout, size=(600,300))

while True:
    b, v = main_window.Read(timeout=100)

    if b in (None, 'Exit'):
        break

    final_list = []

    if b == "folderButton":
        folder = Sg.PopupGetFolder("Select folder")
        folderloc = main_window.FindElement("folderloc")
        folderloc.Update("{}".format(folder))

    if b == "run":

        prog = main_window['progressbar']

        if folder is "":
            Sg.PopupOK("You need to select a folder :)")
            continue

        pages = map(int, v['ranged'].split(','))

        for i in pages:
            final_list.append(i-1)
            print("Appended: ", i-1, " to pages list")

        totalfiles = listdir(folder)
        totalfiles2 = len(totalfiles)
        prog.UpdateBar(1, max=totalfiles2)

        fileloopcount = 0
        totalerrors = 0
        skippedpages = 0
        for x in listdir(folder):
            fileloopcount = fileloopcount + 1
            try:
                hi = PyPDF2.PdfFileReader('{}/{}'.format(folder, x), 'rb')
            except (IndexError, PyPdfError, PdfReadWarning, PdfReadError,
                    PageSizeNotDefinedError, PdfStreamError) as Error:
                Sg.PopupOK("Error with file: {}, Details:\n {}".format(x, Error))
                totalerrors = totalerrors + 1
                continue

            bye = PyPDF2.PdfFileWriter()

            for i in range(0, 69):
                if i in final_list:
                    skippedpages = skippedpages + 1
                    pass
                else:
                    try:
                        page = hi.getPage(i)
                        bye.addPage(page)
                    except (IndexError, PyPdfError, PdfReadWarning, PdfReadError,
                    PageSizeNotDefinedError, PdfStreamError) as Error:
                        if IndexError:
                            continue
                        Sg.PopupOK("Error with file: {}, Details:\n {"
                                   "}".format(x, Error))
                        totalerrors = totalerrors + 1





            with open('{}/{}'.format(folder, x), 'wb') as f:
                bye.write(f)

            prog.UpdateBar(fileloopcount, totalfiles2)

        Sg.PopupOK("All Done\nScanned {} files, {} errors, {} removed pages"
                   ".".format(
                fileloopcount, totalerrors, skippedpages))
        fileloopcount = 0
        totalerrors = 0
        skippedpages = 0

    if b is "Exit" or b is None:
        quit()


main_window.Close()
