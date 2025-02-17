import pythoncom
from win32com.client import Dispatch, gencache
import os

def main():

    # Подключаем API интерфейсов
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    kompas_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)

    # Подключаем объекты верхнего уровня
    application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch)) #IApplication
    kompas_object = kompas_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

    # Получаем текущий чертёж, в котором находится графика для подготовки DXF
    drawing_doc = application.ActiveDocument # IKompasDocument
    drawing_2d1 = kompas_api7_module.IKompasDocument2D1(drawing_doc) # IKompasDocument2D1

    # Создаём группу и помещаем её в буфер обмена
    drawing_selection = drawing_2d1.SelectionManager
    drawing_group = drawing_2d1.DrawingGroups.Add(True, 'copy')
    drawing_group.AddObjects(drawing_selection.SelectedObjects)
    drawing_group.WriteToClip(False, False)


    # Получаем путь к текущему чертежу
    drawing_folder_path = drawing_doc.Path
    drawing_name = os.path.splitext(os.path.basename(drawing_doc.PathName))[0]


    # Создаём новый фрагмент
    documents = application.Documents
    fragment_doc = documents.AddWithDefaultSettings(kompas6_constants.ksDocumentFragment, True) # IKompasDocument
    fragment_2d = kompas_api7_module.IKompasDocument2D(fragment_doc) # IKompasDocument2D
    fragment_2d1 = kompas_api7_module.IKompasDocument2D1(fragment_doc) # IKompasDocument2D1

    # Вставляем скопированную графику
    fragment_group = fragment_2d1.DrawingGroups.Add(True, 'paste')
    fragment_group.ReadFromClip(False, False)
    # fragment_group.Draw(fragment_doc.DocumentFrames.Item(0).GetHWND())
    fragment_group.Store()

    # Сохраняем фрагмент как DXF
    ks_fragment = kompas_object.ActiveDocument2D()
    ks_fragment.ksSaveToDXF(drawing_folder_path + '\\' + drawing_name + '.dxf')
    fragment_doc.Close(0)
    pass




if __name__ == "__main__":
    main()