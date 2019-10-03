# -*- coding: utf_8 -*-
"""
Вставка внизу страницы (во врезку) первого слова со следующей страницы.
Если за словом неразрывный пробел, то первых двух.
Начальные и конечные пробелы и знаки препинания (кроме конечной точки) удаляются.
Сохраняется некоторое форматирование (цвет, и в случае стиля для ч/б печати - жирность)
Настроенные стили: врезки и содержимого врезки (абзацный) предполагается в наличии,
при отсутствии создаются автоматически, и отчасти настраиваются.
Далее их можно настроить под свои нужды.

Запускается из LOffice для открытого документа.
"""

import uno
# import unohelper
from screen_io import MsgBox, InputBox, Print
# from com.sun.star.lang import IndexOutOfBoundsException

frame_prefix = "FWFrame_"
# настроенный cтиль врезки
frame_style_name = "ВрезкаСловоСледСтр"
# настроенный абзацный cтиль для содержимого врезки
frame_paragaph_style_name = "Содержимое врезки ПервоеСловоСледСтр"
# стиль символов
char_style_name = "киноварь"


context = XSCRIPTCONTEXT
desktop = context.getDesktop()
doc = desktop.getCurrentComponent()

style_families = doc.getStyleFamilies()
char_styles = style_families.getByName("CharacterStyles")
# Если есть стиль киноварь, получить значение его цвета.
# В дальнейшем, если он не красный (в стиле для ч/б печати),
# будет учитываться "жирность" при вставке текста во врезку.
if char_styles.hasByName(char_style_name):
    kinovar_color = char_styles.getByName(char_style_name).CharColor
else:
    kinovar_color = 0


def remove_all(*args):
    view_data = save_pos()
    remove_first_words_frames()
    restore_pos_from(view_data)


def update_all(*args):
    view_data = save_pos()
    remove_first_words_frames()
    insert_frames_to_pages()
    restore_pos_from(view_data)


def insert_fw_to_doc(*args):
    """
    В конце каждой страницы вставляет первое слово след-й страницы

    """
    # Если нет стиля врезки, создать.
    check_and_create_styles()

    # Сохранить положение видимого курсора.
    view_data = save_pos()
    # Создать\заполнить врезки.
    insert_frames_to_pages()
    # Вернуть курсор в первоначальное положение.
    restore_pos_from(view_data)


def insert_frames_to_pages():
    """Вставляет в каждую страницу документа фрейм (и заполняет его).

    NOTE: м.б. заполнить отдельно?
    """
    n_pages = doc.getCurrentController().PageCount

    if n_pages == 1:
        return None

    # Получить список курсоров с первыми словами
    # 0-й элемент - первое слово 2-й страницы
    # последний - первое слово последней страницы
    # a = init_fw_array()
    a = [get_fw_from(page) for page in range(2, n_pages + 1)]
    if not a:
        return None

    try:
        # для всех страниц, кроме последней
        for page in range(1, n_pages):
            fw_cursor = a[page - 1]  # курсор с первым словом след-й сраницы
            frame = create_frame_on(page)  # создать (или получить имеющуюся) врезку
            # Если есть врезка и есть слово на след-й странице и оно изменилось,
            if frame and fw_cursor and fw_cursor.getString() != frame.getString():
                fill_frame(frame, fw_cursor)  # занести слово во врезку
                # Если отличие ТОЛЬКО в тексте, то при различии в оформлении (цвет),
                # необходимо вручную удалить и создать заново,
                # (или, что то же, макросом полного обновления всех врезок)

    except IndexError:
        Print("Index Error!")
        return None


# def init_fw_array():
#     """Возвращает список курсоров (для сохранения формата) с первыми словами каждой страницы, начиная со второй.
#
#     :return:
#     """
#     pages_in_doc = doc.getCurrentController().PageCount
#     return [get_fw_from(page) for page in range(2, pages_in_doc + 1)]


def bound_handler(string: str, bound_type='') -> int:
    """Работа с границами цся Unicode-словами в локали "cu".

    Обычным способом (goToEndOfWord) некорректно определяется верхняя граница ЦСЯ слова,
    в случае окончания такими символами, как вария и буквенные титла.

    :param string: Строка с одним или более словами.
    :param bound_type: Тип границы [start, end, next_start, next_end].
    :return: Позиция границ первого или второго слова в строке (0-based).
    """
    from com.sun.star.i18n.WordType import WORD_COUNT, DICTIONARY_WORD
    ctx = uno.getComponentContext()

    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)

    nextwd_bound = uno.createUnoStruct("com.sun.star.i18n.Boundary")
    firstwd_bound = uno.createUnoStruct("com.sun.star.i18n.Boundary")
    a_locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    a_locale.Language = "cu"
    a_locale.Country = "RU"
    mystartpos = 0  # начальная позиция
    brk = create("com.sun.star.i18n.BreakIterator")

    nextwd_bound = brk.nextWord(string, mystartpos, a_locale, DICTIONARY_WORD)
    firstwd_bound = brk.previousWord(string, nextwd_bound.startPos, a_locale, DICTIONARY_WORD)

    bound_dic = {'start': firstwd_bound.startPos,
                 'end': firstwd_bound.endPos,
                 'next_start': nextwd_bound.startPos,
                 'next_end': nextwd_bound.endPos
                 }

    return bound_dic.get(bound_type, -1)


def get_bound_start_pos(string: str) -> int:
    # возвращает позицию нижней границы первого слова в строке
    return bound_handler(string, 'start')


def get_bound_end_pos(string) -> int:
    # возвращает позицию верхней границы первого слова в строке
    return bound_handler(string, 'end')


def get_next_bound_start_pos(string) -> int:
    # возвращает позицию нижней границы второго слова в строке
    return bound_handler(string, 'next_start')


def get_next_bound_end_pos(string) -> int:
    # возвращает позицию верхней границы второго слова в строке
    return bound_handler(string, 'next_end')


def get_fw_from(page: int):
    """На текущей странице сохраняет первое слово.

    Два первых слова, соединенных неразрывным пробелом, рассматриваются как одно.

    :param page: Текущая страница
    :return: Курсор с первым(-и) словом(-ами)
    """
    view_cursor = doc.getCurrentController().getViewCursor()

    view_cursor.jumpToPage(page)
    view_cursor.jumpToStartOfPage()

    # текстовые курсоры
    tmp_cursor = doc.Text.createTextCursorByRange(view_cursor)  # для перемещения
    out_cursor = doc.Text.createTextCursorByRange(view_cursor)  # для конечного захвата

    # захватить одно слово+пробелы и пр. в курсор
    prefix = ''
    first_string = ''  # первое слово с хвостом
    bound_end = -1  # верхняя граница первого слова
    i = 0
    while bound_end < 0:
        # если что-то пойдет не так
        i += 1
        if i >= 30:
            Print("Error on bounds of word!")
            break

        prefix: str = tmp_cursor.getString()  # сохранить для дальнейшего анализа
        tmp_cursor.collapseToEnd()
        tmp_cursor.gotoNextWord(True)
        # -> к началу след-го слова. Если перед первым словом был, к примеру, пробел, или знак тысячи,
        # то курсор перейдет к началу первого слова. Потребуется еще один шаг.
        # Все, находящееся непосредственно перед первым словом, попадет в префикс.

        # На случай пустой страницы. gotoNextWord() может перейти и на след-ю страницу.
        view_cursor.gotoRange(tmp_cursor.getStart(), False)
        vc_current_page = view_cursor.getPage()  # текущая страница view_cursor
        if vc_current_page != page:  # если ушли с текущей странцы
            return None

        # NOTE: следующее слово может также потребовать проверки хвоста (перевод строки и т.п.)
        first_string: str = tmp_cursor.getString()  # вся строка = первое слово + хвост
        bound_end = get_bound_end_pos(first_string)  # верхняя граница первого слова (нижнняя = 0)

    if bound_end > 0:
        founded_first_word = first_string[:bound_end]  # Найденое первое слово
        tail = first_string[bound_end:]  # его хвост

        # Поместить выводящий курсор в начало первого слова
        out_cursor.gotoRange(tmp_cursor.getEnd(), False)  # -> начало след-го слова
        out_cursor.collapseToEnd()
        out_cursor.goLeft(len(tail) + len(founded_first_word), False)  # начало 1-го слова

        # Если в префиксе знак тысячи (или что-то еще нужное)
        capture_end = bound_end  # верняя позиция для захвата курсором
        add_shift = 0  # добавочный сдвиг для случая с неразр.пробелом
        if prefix and prefix[-1] == '҂':
            # скорректировать позиции захвата
            out_cursor.goLeft(1, False)
            capture_end += 1
            add_shift += 1
        out_cursor.collapseToStart()  # -> в начальную позицию захвата

        # В случае с неразравным пробелом
        if first_string.count('\xA0') and first_string[-1] != " ":
            tmp_cursor.gotoNextWord(True)  # в курсоре - два слова с хвостами
            string_with_both_words: str = tmp_cursor.getString()
            bound_next_end = get_next_bound_end_pos(string_with_both_words)
            capture_end = bound_next_end + add_shift  # новая верхняя граница для захвата

        # Проверка на самый крайний случай, чтобы во врезку не попал слишком большой текст
        capture_limit = 35  # ~ половина строки
        if capture_end > capture_limit:
            capture_end = capture_limit

        out_cursor.goRight(capture_end, True)  # захват от начала первого слова до конца блока.
        return out_cursor
    else:
        return None


def create_frame_on(page: int):
    """Возвращает либо имеющуюся, либо новосозданную врезку на странице.

    Врезка именуется "frame_prefix_N" N - номер страницы.

    :param page: Номер страницы.
    :return: врезка на странице.
    """
    view_cursor = doc.getCurrentController().getViewCursor()
    frame_name = frame_prefix + str(page)

    # проверка на сущ-е фрейма
    frames_in_doc = doc.getTextFrames()
    if frames_in_doc.hasByName(frame_name):
        return frames_in_doc.getByName(frame_name)
    else:
        # создание фрейма, если его нет
        frame = doc.createInstance("com.sun.star.text.TextFrame")
        frame.AnchorType = 2  # тип привязки
        frame.FrameStyleName = frame_style_name  # настроенный стиль
        frame.Name = frame_name
        frame.WidthType = 2  # 2 - auto-width
        frame.AnchorPageNo = page  # точное указание страницы для врезки.
        # Решает проблему многостраничного абзаца, когда все врезки на этих страницах
        # помещаются на страницу начала такого абзаца.

        # вставка фрейма в позицию курсора
        view_cursor.jumpToPage(page)
        doc.Text.insertTextContent(view_cursor, frame, 'false')

        return frame


def fill_frame(frame, fw_cursor):
    """
    Вставляет во врезку текст из курсора вместе с форматом (цвет и возможно жирность).

    :param frame: врезка
    :param fw_cursor: курсор с первым словом
    :return:
    """
    # Очистка текста фрейма, т.к. запись нужна
    # либо в новый, либо в устаревший фрейм
    frame.String = ""

    # временный курсор -> во врезку
    tmp_cursor = frame.createTextCursorByRange(frame.getStart())
    # применить абзацный стиль для содержимого врезки
    tmp_cursor.ParaStyleName = frame_paragaph_style_name

    # Структура для сохранения форматирования
    char_props = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
        uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    )

    # Получить текст и формат для всех порций текста, и вставить во врезку
    word_enum = fw_cursor.createEnumeration()  # SwXParagraphEnumeration
    while word_enum.hasMoreElements():
        word = word_enum.nextElement()  # SwXParagraph
        part_of_word_enum = word.createEnumeration()  # SwXTextPortionEnumeration
        while part_of_word_enum.hasMoreElements():
            part_of_word = part_of_word_enum.nextElement()  # SwXTextPortion
            # свойства символов порции
            char_props[0].Name = "CharColor"
            char_props[0].Value = part_of_word.CharColor
            char_props[1].Name = "CharWeight"
            # если есть стиль для цветной печати, то жирность не нужна
            if kinovar_color == 0:
                char_props[1].Value = part_of_word.CharWeight  # если нужен bold
            else:
                char_props[1].Value = 100  # выделения цветом достаточно
            # текст порции
            text_of_part_of_word = part_of_word.String

            # Вставка порции текста во врезку с сохранением формата, через tmp_cursor
            frame.insertTextPortion(text_of_part_of_word, char_props, tmp_cursor)
            tmp_cursor.gotoEndOfWord(False)  # позиция для следующей порции

    return None


def check_and_create_styles():
    """
    Проверка, если нет стиля врезки, и для его содержимого, создает их.
    Далее эти стили можно настроить под свои нужды

    """

    # _style_families = doc.getStyleFamilies()
    frame_styles = style_families.getByName("FrameStyles")
    para_styles = style_families.getByName("ParagraphStyles")

    # Проверка, если нет стиля для врезки, создать
    if not frame_styles.hasByName(frame_style_name):
        MsgBox(f'Нет настроенного стиля врезки. Создаем.')
        new_frame_style = doc.createInstance("com.sun.star.style.FrameStyle")
        new_frame_style.setName(frame_style_name)
        new_frame_style.AnchorType = 2  # AT_PAGE
        new_frame_style.BorderDistance = 0
        new_frame_style.BottomBorderDistance = 0
        new_frame_style.BottomMargin = 1300
        new_frame_style.HoriOrient = 1
        new_frame_style.HoriOrientRelation = 8
        new_frame_style.IsFollowingTextFlow = 'True'
        new_frame_style.LeftBorderDistance = 0
        new_frame_style.LeftMargin = 0
        new_frame_style.PositionProtected = 'True'  # ????
        new_frame_style.RightBorderDistance = 0
        new_frame_style.RightMargin = 0
        new_frame_style.Surround = 1  # PARALLEL
        new_frame_style.TextVerticalAdjust = 1  # TOP
        new_frame_style.TextWrap = 1  # PARALLEL
        new_frame_style.TopBorderDistance = 0
        new_frame_style.VertOrient = 3
        new_frame_style.VertOrientRelation = 7
        new_frame_style.WidthType = 2

        frame_styles.insertByName(frame_style_name, new_frame_style)
        if frame_styles.hasByName(frame_style_name):
            frame_styles.getByName(frame_style_name).ParentStyle = 'Frame'
            MsgBox(f'Cтиль врезки:\n"{frame_style_name}"\nсоздан.')

    # Проверка, если нет стиля для содержимого врезки, создать
    if not para_styles.hasByName(frame_paragaph_style_name):
        new_para_style = doc.createInstance("com.sun.star.style.ParagraphStyle")

        new_para_style.setName(frame_paragaph_style_name)
        new_para_style.ParaAdjust = 1
        new_para_style.CharNoHyphenation = 'True'
        # new_para_style.ParaHyphenationMaxHyphens = 3
        # new_para_style.ParaHyphenationMaxLeadingChars = 4
        # new_para_style.ParaHyphenationMaxTrailingChars = 4
        new_para_style.ParaOrphans = 0
        new_para_style.ParaWidows = 0

        para_styles.insertByName(frame_paragaph_style_name, new_para_style)

        if para_styles.hasByName(frame_paragaph_style_name):
            para_styles.getByName(frame_paragaph_style_name).ParentStyle = "Frame contents"
            MsgBox(f'Cтиль абзаца для содержимого врезки:\n"{frame_paragaph_style_name}"\nсоздан.')


def remove_first_words_frames():
    # Удаляет все врезки созданные для первого слова

    all_frames = doc.getTextFrames()
    frame_names = all_frames.getElementNames()
    for name in frame_names:
        # удалить все врезки вида "frame_prefix_page"
        if name.startswith(frame_prefix):
            all_frames.getByName(name).dispose()


def save_pos():
    # возвращает сохраненную позицию
    return doc.getCurrentController().getViewData()


def restore_pos_from(saved_view_data):
    # восстанавливает позицию
    doc.getCurrentController().restoreViewData(saved_view_data)
    return None


g_exportedScripts = insert_fw_to_doc, remove_all, update_all,
