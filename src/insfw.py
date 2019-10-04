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
import re
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
THOUSAND = '҂'

context = XSCRIPTCONTEXT
desktop = context.getDesktop()
doc = desktop.getCurrentComponent()
n_pages = doc.getCurrentController().PageCount

style_families = doc.getStyleFamilies()
char_styles = style_families.getByName("CharacterStyles")
# Если есть стиль киноварь, получить значение его цвета.
# В дальнейшем, если он не красный (в стиле для ч/б печати),
# будет учитываться "жирность" при вставке текста во врезку.
if char_styles.hasByName(char_style_name):
    kinovar_color = char_styles.getByName(char_style_name).CharColor
else:
    kinovar_color = 0


def Mri_test():
    ctx = context.getComponentContext()
    document = context.getDocument()
    mri(ctx, document)


def Mri(target):
    ctx = context.getComponentContext()
    mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri", ctx)
    mri.inspect(target)


def mri(ctx, target):
    mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri",ctx)
    mri.inspect(target)


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

    """

    if n_pages == 1:
        return None

    # позиции начала и конца каждой страницы (кроме первой)
    pages_positions = [get_start_end_positions_of(page) for page in range(2, n_pages + 1)]

    # Список курсоров с первыми словами для всех страниц, кроме первой
    cursors_with_fword = get_fw_cursors(pages_positions)

    # Врезки с первой по предпоследнюю страницу,
    frames = create_frames_in_doc()

    if not cursors_with_fword or not frames:
        return None

    # кол-во врезок и курсоров должно совпадать с n-1 страниц в док-те.
    if not (n_pages-1 == len(frames) == len(cursors_with_fword)):
        return None

    for frame, cursor in zip(frames, cursors_with_fword):
        # Если есть последнее слово и врезка
        if frame and cursor:  # and frame.String != cursor.String:
            fill_frame(frame, cursor)  # занести слово во врезку


def create_frames_in_doc():
    out = []
    # для всех страниц, кроме последней
    for page in range(1, n_pages):
        frame = create_frame_on(page)  # создать (или получить имеющуюся) врезку
        if frame:
            out.append(frame)
        else:
            out.append(None)

    return out


def get_fw_cursors(pages_positions):
    out = []
    for start, end in pages_positions:
        comparing = doc.Text.compareRegionStarts(start, end)
        if comparing:
            text_cursor = doc.Text.createTextCursorByRange(start)
            text_cursor.gotoRange(start, False)
            text_cursor.gotoRange(end, True)
            fw_cursor = get_fist_word_from_one(text_cursor)
            if fw_cursor:
                out.append(fw_cursor)
            else:
                out.append(None)
        else:
            out.append(None)
    return out


def get_fist_word_from_one(cursor):
    """
    Из курсора с текстом страницы выбирает первое слово (или два)

    :param cursor: курсор с текстом
    :return: курсор с первым словом (или пустой)
    """
    if cursor:
        # Текстовый курсор, который захватит слово с форматом.
        out_cursor = doc.Text.createTextCursorByRange(cursor.getStart())
        page_text = cursor.getString()  # текст всей страницы.

        # Если страница без текста, но с пробелами и пустыми строками,
        # то это предохранит от лишних движений и возможно ошибок.
        page_text = re.sub(r'^\s*$', '', page_text)
        # page_text = re.sub(r'^\s+', '', page_text)

        start_sentence_pos = bound_handler(page_text, 'start_sentence')
        # Если нашлось предложение (не факт, что далее будет именно слово)
        if start_sentence_pos >= 0:
            # Двигаться по тексту, пока не найдется конец слова (первого)
            first_word_start_pos = -1
            first_word_end_pos = -1
            second_word_start_pos = -1
            tmp_position = start_sentence_pos
            i = 0
            while first_word_end_pos == -1:
                i += 1
                first_word_end_pos = get_bound_end_pos(page_text, tmp_position)
                first_word_start_pos = get_bound_start_pos(page_text, tmp_position)
                second_word_start_pos = get_next_bound_start_pos(page_text, tmp_position)
                start_sentence_pos = bound_handler(page_text, 'start_sentence', tmp_position)
                tmp_position = second_word_start_pos

                # MsgBox(
                #     f'Шаг {i}. tmp: {tmp_position}\nsentence: {start_sentence_pos}\n'
                #     f'1-е слово: [{first_word_start_pos}:{first_word_end_pos}] '
                #     f'|{page_text[first_word_start_pos:first_word_end_pos]}|\n'
                #     f'2-е слово: [{second_word_start_pos}:]\n'
                #     # f'tmp_sent = {tmp_sent}'
                # )

                # Если нет приемлемого текста, но страница не совсем пуста.
                if (
                        second_word_start_pos > 0
                        and first_word_start_pos == -1
                        and first_word_end_pos == -1
                ):
                    # MsgBox("не нашлось слова")
                    break

                if i > 100:
                    # MsgBox("не нашлось слова")
                    break  # во избежание зависания
            else:
                # Найден конец первого слова.
                second_word_end_pos = get_next_bound_end_pos(page_text, first_word_start_pos)
                first_word = page_text[first_word_start_pos:first_word_end_pos]
                # промежуток между первым и вторым словом
                tail = page_text[first_word_end_pos:second_word_start_pos]
                # Префикс - между началом предложения и началом слова.
                prefix = page_text[start_sentence_pos:first_word_start_pos]
                # Полный префикс от начала страницы
                full_prefix = page_text[:first_word_start_pos]
                second_word = page_text[second_word_start_pos:second_word_end_pos]

                capture_start_pos = len(full_prefix)  # начальная позиция захвата
                capture_amount = len(first_word)  # захват только 1-го слова

                # Если между первыми двумя словами есть пробел или табуляция,
                # то берется только 1-е слово,
                # иначе (напр. если слова соединены неразр. пробелом), два.
                if not (tail.count(' ') or tail.count('\t')):
                    capture_amount += len(tail + second_word)

                #  Если в префиксе был знак тысячи,
                #  скорректировать позицию и количество захвата
                if prefix and prefix[-1] == THOUSAND:
                    capture_amount += 1
                    capture_start_pos -= 1

                out_cursor.goRight(capture_start_pos, False)  # -> в позицию захвата
                out_cursor.goRight(capture_amount, True)  # захват

    else:
        # Не нашлось слова, но заносим None
        # для соответствия врезкам и страницам
        out_cursor = None

    return out_cursor


def get_start_end_positions_of(page: int):
    """
    Возвращает TextRange - позиции начала и конца страницы.

    (Для получения текста всей страницы и вставки врезки).
    :param page: номер страницы
    :return: tuple of TextRange
    """
    view_cursor = doc.getCurrentController().getViewCursor()
    view_cursor.jumpToPage(page)
    view_cursor.jumpToStartOfPage()
    start_pos = view_cursor.getStart()
    view_cursor.jumpToEndOfPage()
    end_pos = view_cursor.getEnd()
    return start_pos, end_pos


def bound_handler(string: str, bound_type='', start_position=0) -> int:
    """Работа с границами цся Unicode-словами в локали "cu".

    Обычным способом (goToEndOfWord) некорректно определяется верхняя граница ЦСЯ слова,
    в случае окончания такими символами, как вария и буквенные титла.

    :param string: Строка с одним или более словами.
    :param bound_type: Тип границы [start, end, next_start, next_end].
    :param start_position: позиция начала поиска
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
    # mystartpos = 0  # начальная позиция
    brk = create("com.sun.star.i18n.BreakIterator")

    nextwd_bound = brk.nextWord(string, start_position, a_locale, DICTIONARY_WORD)
    firstwd_bound = brk.previousWord(string, nextwd_bound.startPos, a_locale, DICTIONARY_WORD)
    start_of_sentence = brk.beginOfSentence(string, start_position, a_locale)

    bound_dic = {'start': firstwd_bound.startPos,
                 'end': firstwd_bound.endPos,
                 'next_start': nextwd_bound.startPos,
                 'next_end': nextwd_bound.endPos,
                 'start_sentence': start_of_sentence,
                 }

    return bound_dic.get(bound_type, -1)


def get_bound_start_pos(string: str, pos=0) -> int:
    # возвращает позицию нижней границы первого слова в строке
    return bound_handler(string, 'start', pos)


def get_bound_end_pos(string, pos=0) -> int:
    # возвращает позицию верхней границы первого слова в строке
    return bound_handler(string, 'end', pos)


def get_next_bound_start_pos(string, pos=0) -> int:
    # возвращает позицию нижней границы второго слова в строке
    return bound_handler(string, 'next_start', pos)


def get_next_bound_end_pos(string, pos=0) -> int:
    # возвращает позицию верхней границы второго слова в строке
    return bound_handler(string, 'next_end', pos)


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
        MsgBox(f'Нет настроенного стиля для содержимого врезки. Создаем.')
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
