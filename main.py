from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor, Pt

def run_get_spacing(run):
    rPr = run._r.get_or_add_rPr()
    spacings = rPr.xpath("./w:spacing")
    return spacings

def run_get_scale(run):
    rPr = run._r.get_or_add_rPr()
    scale = rPr.xpath("./w:w")
    return scale

def extract_message_from_docx(file_path):
    doc = Document(file_path)
    method = 'Ошибка...'
    extracted_message = ""
    Letters = ''


    for para in doc.paragraphs:
        for run in para.runs:
            char_value = None

            # Проверка цвета текста
            #if run.font.color.rgb != RGBColor(0, 0, 0):
            #    char_value = "1"
            #    method = 'Цвет текста'
            #    Letters += run.text

            # Проверка цвета фона
            #elif run.font.highlight_color != WD_COLOR_INDEX.WHITE:
            #    char_value = "1"
            #    method = 'Цвет фона'
            #    Letters += run.text

            # Проверка размера шрифта
            if run.font.size and run.font.size == Pt(30):
                continue

            if run.font.size and run.font.size != Pt(15):
                char_value = "1"
                method = 'Размер шрифта'
                Letters += run.text

            # Проверка масштаба шрифта
            #elif run_get_scale(run):
            #    char_value = "1"
            #    method = 'Масштаб шрифта'
            #    Letters += run.text

            # Проверка межсимвольного интервала
            #elif run_get_spacing(run):
            #    char_value = "1"
            #    method = 'Межсимвольный интервал'
            #    Letters += run.text

            else:
                char_value = "0"

            # Если значение было определено, добавляем его столько раз, сколько символов в run
            if char_value is not None:
                extracted_message += char_value * len(run.text)

    return extracted_message, method, Letters

file_path = 'Итог.docx'
message, method, letters = extract_message_from_docx(file_path)[0].rstrip('0'), extract_message_from_docx(file_path)[1], extract_message_from_docx(file_path)[2]
print('Метод стеганографического сокрытия:', method)
print('Символ целиком:', letters)

message += '0'*(8-len(message)%8)
byte_data = int(message, 2).to_bytes((len(message) + 7) // 8, byteorder='big')

print('Формирование с помощью 0 и 1:', message)

try:
    decoded_koi8r = byte_data.decode('koi8-r')
    print('КОИ-8R:', decoded_koi8r)
except UnicodeDecodeError:
    print('Ошибка декодирования КОИ-8R')

try:
    decoded_cp866 = byte_data.decode('cp866')
    print('cp866:', decoded_cp866)
except UnicodeDecodeError:
    print('Ошибка декодирования cp866')

try:
    decoded_win1251 = byte_data.decode('windows-1251')
    print('Windows 1251:', decoded_win1251)
except UnicodeDecodeError:
    print('Ошибка декодирования Windows 1251')

try:
    message += '0' * (5 - len(message) % 5)

    bodo_table_lat = {
        '00011': 'A', '11001': 'B', '01110': 'C', '01001': 'D', '00001': 'E',
        '01101': 'F', '11010': 'G', '10100': 'H', '00110': 'I', '01011': 'J',
        '01111': 'K', '10010': 'L', '11100': 'M', '01100': 'N', '11000': 'O',
        '10110': 'P', '10111': 'Q', '01010': 'R', '00101': 'S', '10000': 'T',
        '00111': 'U', '11110': 'V', '10011': 'W', '11101': 'X', '10101': 'Y',
        '10001': 'Z',
        '00010': '\n',
        '00100': ' '
    }
    bodo_table_ru = {
        '00011': 'А', '11001': 'Б', '01110': 'Ц', '01001': 'Д', '00001': 'Е',
        '01101': 'Ф', '11010': 'Г', '10100': 'Х', '00110': 'И', '01011': 'Й',
        '01111': 'К', '10010': 'Л', '11100': 'М', '01100': 'Н', '11000': 'О',
        '10110': 'П', '10111': 'Я', '01010': 'Р', '00101': 'С', '10000': 'Т',
        '00111': 'У', '11110': 'Ж', '10011': 'В', '11101': 'Ь', '10101': 'Ы',
        '10001': 'З',
        '00010': '\n',
        '00100': ' '
    }

    bodo_table_digits = {
        '00011': '-', '11001': '?', '01110': ':', '01001': 'Кто там?', '00001': 'З',
        '01101': 'Э', '11010': 'Ш', '10100': 'Щ', '00110': '8', '01011': 'Ю',
        '01111': '(', '10010': ')', '11100': '.', '01100': ',', '11000': '9',
        '10110': '0', '10111': '1', '01010': '4', '00101': '`', '10000': '5',
        '00111': '7', '11110': '=', '10011': '2', '11101': '/', '10101': '6',
        '10001': '+',
        '00010': '\n',
        '00100': ' '
    }

    current_mode = 'rus'  # По умолчанию ставим русский режим
    decoded_message = ''
    i = 0

    while i < len(message):
        current_bits = message[i:i + 5]

        if current_bits == '11111':
            current_mode = 'lat'
            i += 5
            continue
        elif current_bits == '11011':
            current_mode = 'num'
            i += 5
            continue
        elif current_bits == '00000':
            current_mode = 'rus'
            i += 5
            continue
        else:
            if current_mode == 'lat':
                decoded_message += bodo_table_lat.get(current_bits, '?')
                i += 5
                continue
            elif current_mode == 'num':
                decoded_message += bodo_table_digits.get(current_bits, '?')
                i += 5
                continue
            elif current_mode == 'rus':
                decoded_message += bodo_table_ru.get(current_bits, '?')
                i += 5
                continue

    print('Бодо (МТК-2):', decoded_message)


except:
    print('Ошибка декодирования Бодо (МТК-2)')


