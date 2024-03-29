# -*- coding: utf_8 -*-
from __future__ import unicode_literals
"""
Вставка внизу страницы (во врезку) первого слова со следующей страницы.
Если за словом неразрывный пробел (или нечто подобное), то первых двух.
Начальные и конечные пробелы и знаки препинания (кроме конечной точки) удаляются.
Сохраняется некоторое форматирование (цвет, и в случае стиля для ч/б печати - жирность)
Настроенные стили: врезки и содержимого врезки (абзацный) предполагается в наличии,
при отсутствии создаются автоматически, и отчасти настраиваются.
Далее их можно настроить под свои нужды.

Запускается из LOffice для открытого документа.
"""

import uno
import re
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK


# TODO:
# - не ставить врезку на титульной странице.
# - случай если врезки были созданы при цветном стиле и нужно перейти на ч/б стиль.
# - уточнить логику
#   - обновить все (update_all) - создать заново
#   - добавить: обновить только содержимое врезок
#   - добавить: обновить содержимое только текущей врезки
#   - ?? м.б. разделить защиту содержимого и положения врезки
#   - оставив метод защитить "оба" параметра

FRAME_PREFIX = "FWFrame_"
# настроенный cтиль врезки
FRAME_STYLE_NAME = "ВрезкаСловоСледСтр"
# настроенный абзацный cтиль для содержимого врезки
FRAME_PARAGAPH_STYLE_NAME = "Содержимое врезки ПервоеСловоСледСтр"
# стиль символов
CHAR_STYLE_NAME = "киноварь"
THOUSAND = '҂'


def get_current_component():
    _ctx = uno.getComponentContext()
    _smgr = _ctx.getServiceManager()
    _desktop = _smgr.createInstanceWithContext('com.sun.star.frame.Desktop', _ctx)
    _doc = _desktop.getCurrentComponent()
    if _doc:
        return _doc


def get_pages_abs_count():
    view_data = save_pos()

    _count: int = 1
    view_cursor = doc.getCurrentController().getViewCursor()
    view_cursor.jumpToFirstPage()
    while view_cursor.jumpToNextPage():
        _count += 1

    restore_pos_from(view_data)

    return _count


doc = get_current_component()


def MsgBox(message, title=''):
    '''MsgBox'''

    parent_window = doc.CurrentController.Frame.ContainerWindow
    box = parent_window.getToolkit().createMessageBox(parent_window, MESSAGEBOX, BUTTONS_OK, title, message)
    box.execute()
    return None


class Bound:
    Start = 'start'
    End = 'end'
    NextStart = 'next_start'
    NextEnd = 'next_end'
    StartSentence = 'start_sentence'


class Frame:
    # Frame - вспомогательный класс,
    # обертка над frame из OO (Frame.frame_obj)

    def __new__(cls, page):
        frame_name = FRAME_PREFIX + str(page)
        frames = doc.getTextFrames()
        # объект создается только при наличии врезки на странице
        if frames.hasByName(frame_name):
            return object.__new__(cls)
        else:
            MsgBox('На этой странице нет врезки')
            return None

    def __init__(self, page):
        frame_name = FRAME_PREFIX + str(page)
        frames = doc.getTextFrames()

        self.name = frame_name
        self.page = page
        try:
            self.frame_obj = frames.getByName(frame_name)
        except:
            pass

        self.protected = self.is_protected()
        self.string = self.get_string()

    def get_string(self):
        return self.frame_obj.getString()

    def set_string(self, string):
        self.frame_obj.setString(string)

    def clear(self):
        if not self.is_protected():
            self.set_string('')
        else:
            MsgBox('Содержимое врезки защищено!')

    def delete(self):
        if not self.is_protected():
            self.frame_obj.dispose()
        else:
            MsgBox('Содержимое врезки защищено!')

    def move_up(self):
        self.frame_obj.BottomMargin += 50

    def move_down(self):
        self.frame_obj.BottomMargin -= 50

    def is_protected(self):
        return self.frame_obj.ContentProtected

    def protect(self):
        self.frame_obj.ContentProtected = True
        MsgBox(f'Содержимое врезки на стр.{self.page} защищено')

    def unprotect(self):
        self.frame_obj.ContentProtected = False
        MsgBox(f'Содержимое врезки на стр.{self.page} разблокировано')

    def update_only_current(self):
        # Если это не последня страница
        curr_page = self.page
        insert_frames_to_pages(curr_page + 1, True)


def Mri_test():
    # ctx = context.getComponentContext()
    ctx = uno.getComponentContext()
    # doc = context.getDocument()
    # doc = get_current_component()
    mri(ctx, doc)


def Mri(target):
    # ctx = context.getComponentContext()
    ctx = uno.getComponentContext()
    _mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri", ctx)
    _mri.inspect(target)


def mri(ctx, target):
    _mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri", ctx)
    _mri.inspect(target)


def get_current_page_number():
    return doc.getCurrentController().getViewCursor().getPage()


def remove_all(*args):
    view_data = save_pos()
    remove_first_words_frames()
    restore_pos_from(view_data)


def update_all(*args):
    view_data = save_pos()
    insert_frames_to_pages()
    restore_pos_from(view_data)


def insert_fw_to_doc(*args):
    """В конце каждой страницы вставляет первое слово след-й страницы

    :return
    """
    # Если нет стиля врезки, создать.
    check_and_create_styles()

    # Сохранить положение видимого курсора.
    view_data = save_pos()
    # Создать\заполнить врезки.
    insert_frames_to_pages()
    # Вернуть курсор в первоначальное положение.
    restore_pos_from(view_data)


def insert_frames_to_pages(start_page=2, one_page_flag=False):
    """Вставляет в каждую страницу документа фрейм (и заполняет его).

    :param start_page: начальная страница для обработки. По умолчанию, начиная со второй.
    :param one_page_flag: флаг обработки только текущей страницы.
    """
    n_pages = get_pages_abs_count()
    if one_page_flag:
        end_page = start_page + 1
    else:
        end_page = n_pages

    if n_pages == 1:
        return None

    # позиции начала и конца каждой страницы (кроме первой)
    pages_positions = [get_start_end_positions_of(page) for page in range(start_page, end_page + 1)]

    # Список курсоров с первыми словами для всех страниц, кроме первой
    cursors_with_fword = get_fw_cursors(pages_positions)

    # Для всего документа
    if not one_page_flag:
        # Врезки с первой по предпоследнюю страницу,
        frames = make_all_frames_in(pages_positions)  # , start_page - 2)
    # Для текущей страницы
    else:
        _curr_page = start_page - 1
        _start, _end = get_start_end_positions_of(_curr_page)
        frames = [make_frame_in_position(_curr_page, _end), None]

    if not cursors_with_fword or not frames:
        return None

    # кол-во врезок и курсоров должно совпадать с n-1 страниц в док-те.
    if not (end_page - start_page + 1 == len(frames) == len(cursors_with_fword)):
        return None

    for frame, cursor in zip(frames, cursors_with_fword):
        # Если есть последнее слово и врезка, и если содержимое врезки не защищено,
        if frame and cursor and not frame.ContentProtected:  # and frame.String != cursor.String:
            fill_frame(frame, cursor)  # занести слово во врезку


def make_all_frames_in(pages_positions, _page=0):
    """Создать (или получить имеющиеся) врезки

    :param pages_positions: список кортежей позиций начала и конца страницы.
    :param _page: номер страницы, с которой начинать счет.
    :return: список врезок.
    """
    out_frames = []
    # для всех концов страниц (кроме последней),
    # page = 0
    for _, end in pages_positions:
        _page += 1
        frame = make_frame_in_position(_page, end)  # создать (или получить имеющуюся) врезку
        if frame:
            out_frames.append(frame)
        else:
            out_frames.append(None)

    return out_frames


def get_fw_cursors(pages_positions):
    out = []
    for start, end in pages_positions:
        comparing = doc.Text.compareRegionStarts(start, end)
        # Если страница не пуста
        if comparing:
            # захватить в курсор весь текст
            text_cursor = doc.Text.createTextCursorByRange(start)
            text_cursor.gotoRange(start, False)
            text_cursor.gotoRange(end, True)
            # получить из этого курсора другой, с первым словом
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

    :param cursor: курсор с текстом.
    :return: курсор с первым словом (или пустой).
    """
    if cursor:
        # Текстовый курсор, который захватит слово с форматом.
        out_cursor = doc.Text.createTextCursorByRange(cursor.getStart())
        page_text = cursor.getString()  # текст всей страницы.

        # Если страница без текста, но с пробелами и пустыми строками,
        # то это предохранит от лишних движений и возможно ошибок.
        page_text = re.sub(r'^\s*$', '', page_text)
        # page_text = re.sub(r'^\s+', '', page_text)

        # Если перед первым словом стоят кавычки («),
        # то не работает граница слова.
        page_text = re.sub(r'^«', 'Ѣ', page_text)

        start_sentence_pos = bound_handler(page_text, Bound.StartSentence)
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
                start_sentence_pos = bound_handler(page_text, Bound.StartSentence, tmp_position)
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
                # если слово одно на 1-й строке, но на второй троке есть еще слово,
                # то берется только 1-е слово,
                # иначе (напр. если слова соединены неразр. пробелом), два.
                if not (
                        tail.count(' ')
                        or tail.count('\t')
                        or tail.count('\x0A')  # new line
                        or tail.count('\x0D')  # new para
                ):
                    capture_amount += len(tail + second_word)

                #  Если в префиксе был знак тысячи,
                #  скорректировать позицию и количество захвата
                if prefix and prefix[-1] == THOUSAND:
                    capture_amount += 1
                    capture_start_pos -= 1

                out_cursor.goRight(capture_start_pos, False)  # -> в позицию захвата
                out_cursor.goRight(capture_amount, True)  # захват

    else:
        # Не нашлось слова, но возвращаем None
        # для соответствия по кол-ву и порядку врезкам и страницам
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


def bound_handler(string: str, bound_type: str = None, start_position=0) -> int:
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

    bound_dic = {
        Bound.Start: firstwd_bound.startPos,
        Bound.End: firstwd_bound.endPos,
        Bound.NextStart: nextwd_bound.startPos,
        Bound.NextEnd: nextwd_bound.endPos,
        Bound.StartSentence: start_of_sentence,
    }

    return bound_dic.get(bound_type, -1)


def get_bound_start_pos(string: str, pos=0) -> int:
    # возвращает позицию нижней границы первого слова в строке
    return bound_handler(string, Bound.Start, pos)


def get_bound_end_pos(string, pos=0) -> int:
    # возвращает позицию верхней границы первого слова в строке
    return bound_handler(string, Bound.End, pos)


def get_next_bound_start_pos(string, pos=0) -> int:
    # возвращает позицию нижней границы второго слова в строке
    return bound_handler(string, Bound.NextStart, pos)


def get_next_bound_end_pos(string, pos=0) -> int:
    # возвращает позицию верхней границы второго слова в строке
    return bound_handler(string, Bound.NextEnd, pos)


def make_frame_in_position(page, end_of_page):
    """Возвращает либо имеющуюся, либо новосозданную врезку в позиции конца страницы.

    :param page: номер страницы
    :param end_of_page: позиция конца страницы (TextRange)
    :return: врезка (TextFrame)
    """

    frame_name = FRAME_PREFIX + str(page)
    # проверка на сущ-е фрейма
    frames_in_doc = doc.getTextFrames()
    if frames_in_doc.hasByName(frame_name):
        return frames_in_doc.getByName(frame_name)
    else:
        # создание фрейма, если его нет
        frame = doc.createInstance("com.sun.star.text.TextFrame")
        frame.AnchorType = 2  # тип привязки
        frame.FrameStyleName = FRAME_STYLE_NAME  # настроенный стиль
        frame.Name = frame_name
        frame.WidthType = 2  # 2 - auto-width
        frame.AnchorPageNo = page  # точное указание страницы для врезки.
        # Решает проблему многостраничного абзаца, когда все врезки на этих страницах
        # помещаются на страницу начала такого абзаца.

        # вставка фрейма в позицию конца страницы
        doc.Text.insertTextContent(end_of_page, frame, 'false')
        if frames_in_doc.hasByName(frame_name):
            return frame
        else:
            return None


def fill_frame(frame, fw_cursor):
    """
    Вставляет во врезку текст из курсора вместе с форматом (цвет и возможно жирность).

    :param frame: врезка.
    :param fw_cursor: курсор с первым словом.
    :return:
    """

    char_styles = doc.getStyleFamilies().getByName("CharacterStyles")
    # Если есть стиль киноварь, получить значение его цвета.
    # В дальнейшем, если он не красный (в стиле для ч/б печати),
    # будет учитываться "жирность" при вставке текста во врезку.
    kinovar_color = 0
    if char_styles.hasByName(CHAR_STYLE_NAME):
        kinovar_color = char_styles.getByName(CHAR_STYLE_NAME).CharColor

    # Очистка текста фрейма, т.к. запись нужна
    # либо в новый, либо в устаревший фрейм.
    frame.String = ""

    # временный курсор -> во врезку
    tmp_cursor = frame.createTextCursorByRange(frame.getStart())
    # применить абзацный стиль для содержимого врезки
    tmp_cursor.ParaStyleName = FRAME_PARAGAPH_STYLE_NAME

    # Структура для сохранения форматирования
    char_props = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
        uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    )

    # Получить текст и формат для всех порций текста, и вставить во врезку.
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
            # TODO: оперировать символьными стилями, если есть. Если нет то как обычно
            #  если есть стиль для цветной печати, то жирность не нужна
            #  if kinovar_color in [-1, 0]:
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
    style_families = doc.getStyleFamilies()
    frame_styles = style_families.getByName("FrameStyles")
    para_styles = style_families.getByName("ParagraphStyles")

    # Проверка, если нет стиля для врезки, создать
    if not frame_styles.hasByName(FRAME_STYLE_NAME):
        MsgBox('Нет настроенного стиля врезки. Создаем.')
        new_frame_style = doc.createInstance("com.sun.star.style.FrameStyle")
        new_frame_style.setName(FRAME_STYLE_NAME)
        new_frame_style.AnchorType = 2  # AT_PAGE
        new_frame_style.BorderDistance = 0
        new_frame_style.BottomBorderDistance = 0
        # Для нижнего поля 2300, интерлиньяж одинарный, кегль 20пт.
        new_frame_style.BottomMargin = 1650
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

        frame_styles.insertByName(FRAME_STYLE_NAME, new_frame_style)
        if frame_styles.hasByName(FRAME_STYLE_NAME):
            frame_styles.getByName(FRAME_STYLE_NAME).ParentStyle = 'Frame'
            MsgBox('Cтиль врезки:\n"{}"\nсоздан.'.format(FRAME_STYLE_NAME))

    # Проверка, если нет стиля для содержимого врезки, создать
    if not para_styles.hasByName(FRAME_PARAGAPH_STYLE_NAME):
        MsgBox('Нет настроенного стиля для содержимого врезки. Создаем.')
        new_para_style = doc.createInstance("com.sun.star.style.ParagraphStyle")

        new_para_style.setName(FRAME_PARAGAPH_STYLE_NAME)
        new_para_style.ParaAdjust = 1
        new_para_style.CharNoHyphenation = 'True'
        # new_para_style.ParaHyphenationMaxHyphens = 3
        # new_para_style.ParaHyphenationMaxLeadingChars = 4
        # new_para_style.ParaHyphenationMaxTrailingChars = 4
        new_para_style.ParaOrphans = 0
        new_para_style.ParaWidows = 0

        para_styles.insertByName(FRAME_PARAGAPH_STYLE_NAME, new_para_style)

        if para_styles.hasByName(FRAME_PARAGAPH_STYLE_NAME):
            para_styles.getByName(FRAME_PARAGAPH_STYLE_NAME).ParentStyle = "Frame contents"
            MsgBox('Стиль абзаца для содержимого врезки:\n"{}"\nсоздан.'.format(FRAME_PARAGAPH_STYLE_NAME))


def remove_first_words_frames():
    # Удаляет все врезки созданные для первого слова

    all_frames = doc.getTextFrames()
    frame_names = all_frames.getElementNames()
    for name in frame_names:
        # удалить все врезки вида "frame_prefix_page"
        if name.startswith(FRAME_PREFIX):
            try:
                all_frames.getByName(name).dispose()
            except:
                pass


def save_pos():
    # возвращает сохраненную позицию
    return doc.getCurrentController().getViewData()


def restore_pos_from(saved_view_data):
    # восстанавливает позицию
    doc.getCurrentController().restoreViewData(saved_view_data)
    return None


def update_current_frame(*args):
    # Очищает врезку на текущей странице
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.update_only_current()


def clear_current_frame(*args):
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.clear()


def delete_current_frame(*args):
    # Удаляет врезку на текущей странице
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.delete()


def up_current_frame(*args):
    # поднять врезку на 0.05
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.move_up()


def down_current_frame(*args):
    # опустить врезку на 0.05
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.move_down()


def protect_current_frame(*args):
    # Защитить содержимое врезки
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.protect()


def unprotect_current_frame(*args):
    # Убрать защиту содержимого врезки
    page = get_current_page_number()
    frame = Frame(page)
    if frame:
        frame.unprotect()


g_exportedScripts = (
    insert_fw_to_doc,
    remove_all,
    update_all,
    clear_current_frame,
    down_current_frame,
    up_current_frame,
    protect_current_frame,
    unprotect_current_frame,
    delete_current_frame,
    update_current_frame,
)
