import win32com.client


def calc_block_ref(ncad_document):
    """
    Функция определяет число блоков различных типов в пространстве модели.
    Возвращает словарь, ключомами являются наименования блоков, а значениями - числа вхождения блоков.

    :param ncad_document: файл, в котором необходимо произвести анализ числа вхождения блоков
    :return: словарь, отсортированный по наименованиям блоков
    """
    block2count = dict()
    for model_object in ncad_document.ModelSpace:
        if model_object.EntityName == 'AcDbBlockReference':
            object_as_BlockRef = win32com.client.CastTo(model_object, "IAcadBlockReference")
            if object_as_BlockRef is not None:
                if object_as_BlockRef.Name not in block2count.keys():
                    block2count[object_as_BlockRef.Name] = 1
                else:
                    block2count[object_as_BlockRef.Name] += 1
    print(block2count)
    return block2count


def calc_line_len(ncad_document):
    """
   Функция определяет сумму длин полилиний по слоям в пространстве модели.
   Происходит сортировка полилиний по их принадлежности к слоям и подсчет суммарной длины
   Результатом является словарь, где ключами являются названия слоев, а значениями - сумманые длины линий
   в рамках одного слоя.

   :param ncad_document: файл, в котором необходимо произвести анализ полилиний
   :return: словарь полилиний, отсортированный по наименованиям слоёв
   """
    layer_len_dict = dict()
    for model_object in ncad_document.ModelSpace:
        if model_object.EntityName == 'AcDbPolyline':
            object_as_PolylineRef = win32com.client.CastTo(model_object, "IAcadLWPolyline")
            if object_as_PolylineRef is not None:
                if object_as_PolylineRef.Layer not in layer_len_dict.keys():
                    layer_len_dict[object_as_PolylineRef.Layer] = object_as_PolylineRef.Length
                else:
                    layer_len_dict[object_as_PolylineRef.Layer] += object_as_PolylineRef.Length
    for one_layer in layer_len_dict:
        layer_len_dict[one_layer] = round(layer_len_dict[one_layer], 3)
    return layer_len_dict


def calc_text_symb(ncad_document):
    """
   Функция определяет суммарное количество символов однострочных текстов по слоям в пространстве модели.
   Происходит сортировка однострочного текста по принадлежности к слою и подсчет количества символов на слое.
   Результатом явялется словарь, где ключами являются названия слоев, а значениями - суммарное
   количество текстовых символов в рамках одного слоя.

   :param ncad_document: файл, в котором необходимо произвести анализ полилиний
   :return: словарь, отсортированный по наименованиям слоёв
   """
    layer_symb_dict = dict()
    for model_object in ncad_document.ModelSpace:
        if model_object.EntityName == 'AcDbText':
            object_as_SymbRef = win32com.client.CastTo(model_object, "IAcadText")
            if object_as_SymbRef is not None:
                if object_as_SymbRef.Layer not in layer_symb_dict.keys():
                    layer_symb_dict[object_as_SymbRef.Layer] = len(object_as_SymbRef.TextString)
                else:
                    layer_symb_dict[object_as_SymbRef.Layer] += len(object_as_SymbRef.TextString)
    for one_layer in layer_symb_dict:
        layer_symb_dict[one_layer] = round(layer_symb_dict[one_layer], 3)
    return layer_symb_dict


def calc_hatch_area(ncad_document):
    """
    Функция определяет сумму площадей штриховок по слоям в пространстве модели.
    Происходит сортировка штриховок по принадлежности к слою и подсчет площадей штриховок на слое.
    Результатом является словарь штриховок, где ключами являются названия слоев, а значениями - суммарные площади
    всех штриховок в рамках одного слоя.

    :param ncad_document: layout на котором необходимо произвести анализ полилиний
    :return: словарь, отсортированный по наименованиям слоёв
    """
    layer_hatch_dict = dict()
    for model_object in ncad_document.ModelSpace:
        if model_object.EntityName == 'AcDbHatch':
            object_as_HatchRef = win32com.client.CastTo(model_object, "IAcadHatch")
            if object_as_HatchRef is not None:
                if object_as_HatchRef.Layer not in layer_hatch_dict.keys():
                    layer_hatch_dict[object_as_HatchRef.Layer] = object_as_HatchRef.Area
                else:
                    layer_hatch_dict[object_as_HatchRef.Layer] += object_as_HatchRef.Area
    for one_layer in layer_hatch_dict:
        layer_hatch_dict[one_layer] = round(layer_hatch_dict[one_layer], 3)
    return layer_hatch_dict


start_point = [0.05, 21.78]


def create_table_and_text(
        text_line1: str, text_line2: str, table_title: str, col1_name: str, col2_name: str,
        list_with_obj: list, smth2count: dict
):
    """
    Функция создает блок таблицы и блок MText на на листе "Для вставки таблтиц" в nanoCAD.
    Формируемые таблицы состоят из двух колонок, шириной 36 мм.
    Высота многострочного текста (0.2) должна быть установлена вручную на листе в nanoCAD, т.к. установку
    данного параметра нельзя автоматизировать.
    Также возможно создание нескольких строк однострочного текста, высотой 0,2. Для этого следует раскомментрировать
    соответствующие строки (130-131).

    :param text_line1: Первая строка заголовка таблицы при применении однострочного текста
    :param text_line2: Вторая строка заголовка таблицы при применении однострочного текста
    :param table_title: Текст шапки таблицы
    :param col1_name: наименование первой колонки
    :param col2_name: наименование второй колонки
    :param list_with_obj: список объектов, по которому производилась сортировка элементов
    :param smth2count: печатаемый словарь
    :return: графическое представление таблицы и текста заголовка в nanoCAD
    """
    needed_layout = ncad_doc.Database.Layouts.Item(0)  # выбор листа
    for one_layout in ncad_doc.Database.Layouts:
        if one_layout.Name == "Для вставки таблтиц":
            needed_layout = one_layout
            break

    AcadText_object = needed_layout.Block.AddMText(start_point, 7.2,
                                                   text_line1 + text_line2)  # применение MText без установки высоты текста
    # применение Text с установкой высоты текста и ручной разбивкой на строки
    # AcadText_object = needed_layout.Block.AddText(text_line1, [start_point[0], start_point[1]], 0.2)
    # AcadText_object = needed_layout.Block.AddText(text_line2, [start_point[0], start_point[1]-0.25], 0.2)

    AcadTable_object = needed_layout.Block.AddTable([start_point[0], start_point[1] - 1], len(list_with_obj) + 2, 2,
                                                    0.1, 3.6)
    AcadTable_object.SetText(0, 0, table_title)
    AcadTable_object.SetText(1, 0, col1_name)
    AcadTable_object.SetText(1, 1, col2_name)

    counter_rows = 2
    for one_name in list_with_obj:
        AcadTable_object.SetText(counter_rows, 0, one_name)
        if one_name in smth2count.keys():
            AcadTable_object.SetText(counter_rows, 1, smth2count[one_name])
        else:
            AcadTable_object.SetText(counter_rows, 1, 0)
        counter_rows += 1
    start_point[0] += 7.5
    return


"""
Ниже приведена часть кода, которая подключается к активному окну nanoCAD и выполняет приведенные выше функции, а именно:
Из пространства модели (ModelSpace) получает количественные характеристики:
- количество вхождений блока каждого типа;
- суммарная длина всех линий с сортировкой по слоям;
- суммарное количество текстовых символов во всех Однострочных текстах с сортировкой по слоям;
- суммарная площадь всей штриховки с сортировкой по слоям
Полученные количественные характеристики оформляются в Таблицы и размещаются в пространстве листа "Для вставки таблиц" ,
также к каждой таблице элементом MText создается заголовок. 
"""

nanocad_app = win32com.client.Dispatch("nanoCADx64.Application")
if nanocad_app is not None:
    ncad_doc = nanocad_app.ActiveDocument
    if ncad_doc is not None:
        print('Doc exists ' + ncad_doc.Name)

        # ТАБЛИЦА С КОЛИЧЕСТВОМ ВХОЖДЕНИЙ БЛОКА КАЖДОГО ТИПА
        Blocks = ncad_doc.Database.Blocks
        Blocks_list = list()
        for one_block in Blocks:
            if one_block.Name[0] != '*':
                Blocks_list.append(one_block.Name)
        Blocks_list.sort()

        block2count = calc_block_ref(ncad_doc)  # Считаем вхождения блоков
        create_table_and_text(
            text_line1='Количество вхождений', text_line2='блока каждого типа',
            table_title="Спецификация количества блоков", col1_name="Имя блока", col2_name="Число, шт.",
            list_with_obj=Blocks_list, smth2count=block2count
        )

        # СПИСОК ВСЕХ СЛОЕВ
        Layers = ncad_doc.Database.Layers
        layer_list = list()
        for one_type in Layers:
            layer_list.append(one_type.Name)
        layer_list.sort()

        # СУММАРНАЯ ДЛИНА ВСЕХ ЛИНИЙ С СОРТИРОВКОЙ ПО СЛОЯМ
        lines2count = calc_line_len(ncad_doc)
        print('Сумма длин линий по слоям: ' + str(lines2count))
        create_table_and_text(
            text_line1='Суммарная длина всех линий', text_line2='с сортировкой по слоям',
            table_title="Спецификация длин линий по слоям", col1_name="Имя слоя", col2_name="Длина",
            list_with_obj=layer_list, smth2count=lines2count
        )

        # СУММАРНОЕ КОЛИЧЕСТВО ТЕКСТОВЫХ СИМВОЛОВ ВО ВСЕХ ОДНОСТРОЧНЫХ ТЕКСТАХ С СОРТИРОВКОЙ ПО СЛОЯМ
        symb2count = calc_text_symb(ncad_doc)
        print('Сумма знаков по слоям: ' + str(symb2count))
        create_table_and_text(
            text_line1='Суммарное количество текстовых символов во',
            text_line2='всех однострочных текстах с сортировкой по слоям',
            table_title="Спецификация числа символов по слоям", col1_name="Имя слоя", col2_name="Число символов, шт.",
            list_with_obj=layer_list, smth2count=symb2count
        )

        # СУММАРНАЯ ПЛОЩАДЬ ВСЕЙ ШТРИХОВКИ С СОРТИРОВКОЙ ПО СЛОЯМ
        hatch2count = calc_hatch_area(ncad_doc)
        print('Сумма площадей штриховки по слоям: ' + str(hatch2count))
        create_table_and_text(
            text_line1='Суммарная площадь всей штриховки', text_line2='с сортировкой по слоям',
            table_title="Спецификация площадей штриховки по слоям", col1_name="Имя слоя", col2_name="Площадь",
            list_with_obj=layer_list, smth2count=hatch2count
        )

        print('The end.')
    else:
        print("Doc is not running")
else:
    print("App is not running")
#
