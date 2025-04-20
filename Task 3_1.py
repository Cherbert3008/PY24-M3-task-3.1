import win32com.client

nanocad_app = win32com.client.Dispatch("nanoCAD.Application")
if nanocad_app is not None:
    ncad_doc = nanocad_app.ActiveDocument
    if ncad_doc is not None:

        dic_block = {}
        dic_layers = {}

        for one_layer_index in range(0, ncad_doc.Layers.Count, 1):
            ncad_lay = ncad_doc.Layers.Item(one_layer_index)
            dic_layers[ncad_lay.Name] = {'Длина линий': 0, 'Кол-во текстовых символов': 0, 'Площадь штриховок': 0}

        for one_AcadEntity in ncad_doc.ModelSpace:
            if not one_AcadEntity.ObjectName in dic_block:
                dic_block[one_AcadEntity.ObjectName] = 1
            else:
                dic_block[one_AcadEntity.ObjectName] += 1

            if one_AcadEntity.ObjectName == "AcDbText":
                object_Text = win32com.client.CastTo(one_AcadEntity, "IAcadText")
                dic_layers[str(object_Text.Layer)][ 'Кол-во текстовых символов'] += len(object_Text.TextString)
            else:
                if one_AcadEntity.ObjectName == "AcDbPolyline":
                    object_Polyline = win32com.client.CastTo(one_AcadEntity, "IAcadLWPolyline")
                    dic_layers[str(object_Polyline.Layer)]['Длина линий'] += object_Polyline.Length
                else:
                    if one_AcadEntity.ObjectName == "AcDbHatch":
                        object_Hatch = win32com.client.CastTo(one_AcadEntity, "IAcadHatch")
                        dic_layers[str(object_Hatch.Layer)]['Площадь штриховок'] += object_Hatch.Area

        for one_layout_index in range(0, ncad_doc.Layouts.Count, 1):
            ncad_Layout = ncad_doc.Layouts.Item(one_layout_index)
            if ncad_Layout.Name == "Для вставки таблтиц":
                ncad_Block_for_Layout = ncad_Layout.Block
                size_of_layout = ncad_Layout.GetPaperSize()

                def create_table (ncad_Block_for_Layout, tab_place, dic_table, title, col1, col2, ind=''):
                    Table = ncad_Block_for_Layout.AddTable(tab_place, len(dic_table.keys()) + 2, 2, 5, 150)
                    Table.SetTextHeight(1, 4)
                    Table.SetTextHeight(2, 4)
                    Table.SetTextHeight(4, 4)
                    Table.SetText(0, 0, title)
                    Table.SetText(1, 0, col1)
                    Table.SetText(1, 1, col2)
                    counter_rows = 2
                    if ind == '':
                        for class_name, obj_count in dic_table.items():
                            Table.SetCellAlignment(counter_rows, 0, 7)
                            Table.SetText(counter_rows, 0, class_name)
                            Table.SetText(counter_rows, 1, obj_count)
                            counter_rows += 1
                    else:
                        for class_name, obj_count in dic_table.items():
                            Table.SetCellAlignment(counter_rows, 0, 7)
                            Table.SetText(counter_rows, 0, class_name)
                            Table.SetText(counter_rows, 1, obj_count[ind])
                            counter_rows += 1


                Title_Blocks_Place = str(str(size_of_layout[0] / 2 - 350) + "," + str(size_of_layout[1] - 35) + ",0")
                Table_Blocks_Place = str(str(size_of_layout[0] / 2 - 350) + "," + str(size_of_layout[1] - 50) + ",0")
                Title_Blocks = ncad_Block_for_Layout.AddText('количество Вхождений блока каждого типа', Title_Blocks_Place,5)
                create_table(ncad_Block_for_Layout, Table_Blocks_Place, dic_block, "Спецификация блоков модели", "Объектный класс","Кол-во, шт.")

                Title_Lines_Place = str(str(size_of_layout[0] / 2 - 350) + "," + str(size_of_layout[1] - 105) + ",0")
                Table_Lines_Place = str(str(size_of_layout[0] / 2 - 350) + "," + str(size_of_layout[1] - 120) + ",0")
                Title_Lines = ncad_Block_for_Layout.AddText('суммарная длина всех линий с сортировкой по слоям',Title_Lines_Place, 5)
                create_table(ncad_Block_for_Layout, Table_Lines_Place, dic_layers, "Спецификация полилиний модели","Слой", "Длина полилиний", 'Длина линий')

                Title_Text_Place = str(str(size_of_layout[0] / 2) + "," + str(size_of_layout[1] - 35) + ",0")
                Table_Text_Place = str(str(size_of_layout[0] / 2 ) + "," + str(size_of_layout[1] - 50) + ",0")
                Title_Text = ncad_Block_for_Layout.AddText('суммарное количество текстовых символов во всех Однострочных текстах с сортировкой по слоям',Title_Text_Place, 5)
                create_table(ncad_Block_for_Layout, Table_Text_Place, dic_layers, "Спецификация текста в модели", "Слой", "Кол-во символов", 'Кол-во текстовых символов')

                Title_Hatch_Place = str(str(size_of_layout[0] / 2) + "," + str(size_of_layout[1] - 335) + ",0")
                Table_Hatch_Place = str(str(size_of_layout[0] / 2) + "," + str(size_of_layout[1] - 350) + ",0")
                Title_Hatch = ncad_Block_for_Layout.AddText('суммарная площадь всей штриховки с сортировкой по слоям',Title_Hatch_Place, 5)
                create_table(ncad_Block_for_Layout, Table_Hatch_Place, dic_layers, "Спецификация штриховок модели","Слой", "Площадь", 'Площадь штриховок')

    else:
        print('Doc not running')
else:
    print('App not running')
print ('work is complited')

