# -*- coding: utf_8 -*-
"""
Вставка внизу страницы первого слова со следующей страницы. Слово вставляется во врезку.
Если за словом неразрывный пробел, то первых двух.
Начальные и конечные пробелы и знаки препинания (кроме конечной точки) удаляются.
Сохраняется некоторое форматирование (цвет, и в случае стиля для ч/б печати - жирность)
Настроенный стиль врезки (frame_style_name) предполагается в наличии,
но если его нет то он создается автоматически. Далее его можно настроить под свои нужды.
Также неплохо настроить и стиль абзаца "Содержимое врезки".

Запускается из LOffice для открытого документа.
"""

import uno
# import unohelper
from screen_io import MsgBox, InputBox, Print

frame_style_name = "ВрезкаСловоСледСтр"  # настроенный cтиль врезки
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
    check_and_create_frame_style()

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
    a = init_fw_array()
    if not a:
        return None

    try:
        # для всех станиц, кроме последней
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


def init_fw_array():
    """Возвращает список курсоров (для сохранения формата) с первыми словами каждой страницы.

    :return: list
    """
    pages_in_doc = doc.getCurrentController().PageCount
    if pages_in_doc == 1:
        return None
    a = []
    # начиная со второй страницы
    for page in range(2, pages_in_doc + 1):
        # добавить полученный курсор c первым словом
        a.append(get_fw_from(page))
    return a


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

    if bound_type == 'start':
        return firstwd_bound.startPos
    elif bound_type == 'end':
        return firstwd_bound.endPos
    elif bound_type == 'next_start':
        return nextwd_bound.startPos
    elif bound_type == 'next_end':
        return nextwd_bound.endPos
    else:
        return -1


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
            out_cursor.goLeft(1, False)  # сместиться на 1 символ влево
            capture_end += 1  # скорректировать позиции захвата
            add_shift += 1
        out_cursor.collapseToStart()  # начальная позиция захвата

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

    Врезка именуется "TxtFrame_N" N - номер страницы.

    :param page: Номер страницы.
    :return: врезка на странице.
    """
    from com.sun.star.text.TextContentAnchorType import AT_PAGE
    view_cursor = doc.getCurrentController().getViewCursor()
    frame_name = "TxtFrame_" + str(page)

    # проверка на сущ-е фрейма
    frames_in_doc = doc.getTextFrames()
    if frames_in_doc.hasByName(frame_name):
        return frames_in_doc.getByName(frame_name)
    else:
        # создание фрейма, если его нет
        frame = doc.createInstance("com.sun.star.text.TextFrame")
        frame.AnchorType = AT_PAGE  # тип привязки
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


def check_and_create_frame_style():
    """
    Проверка, если нет стиля для врезки, создает его
    далее этот стиль можно настроить под свои нужды

    """
    _style_families = doc.getStyleFamilies()
    frame_styles = _style_families.getByName("FrameStyles")
    # Проверка, если нет стиля для врезки, создать
    if not frame_styles.hasByName(frame_style_name):
        Print("Нет стиля " + frame_style_name + ", создаем.")
        new_frame_style = doc.createInstance("com.sun.star.style.FrameStyle")
        new_frame_style.setName(frame_style_name)
        frame_styles.insertByName(frame_style_name, new_frame_style)


def remove_first_words_frames():
    # Удаляет все врезки созданные для первого слова

    all_frames = doc.getTextFrames()
    n_pages = doc.getCurrentController().PageCount

    # для каждой страницы кроме последней
    # удалить все врезки вида "TxtFrame_page"
    for page in range(1, n_pages):
        frame_name = "TxtFrame_" + str(page)
        # если есть врезка, удалить ее
        if all_frames.hasByName(frame_name):
            all_frames.getByName(frame_name).dispose()


def save_pos():
    # возвращает сохраненную позицию
    return doc.getCurrentController().getViewData()


def restore_pos_from(saved_view_data):
    # восстанавливает позицию
    doc.getCurrentController().restoreViewData(saved_view_data)
    return None


g_exportedScripts = insert_fw_to_doc, remove_all, update_all,
