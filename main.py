import telebot
import io
import PyPDF2
from docx import Document
from pydub import AudioSegment
import ffmpeg
import easyocr
import os
from docx2pdf import convert
import tempfile

bot = telebot.TeleBot('7136593333:AAEoE6Zx-0fCblBsbQFNCtoVa0LnoyLIYfo')
reader = easyocr.Reader(['en', 'ru'])


@bot.message_handler(commands=['start'])
def main(message):
    bot.send_message(message.chat.id, 'Привет, этот бот способен конвертировать файлы.\n'
    'Для просмотра команд напишите /menu')


@bot.message_handler(commands=['menu'])
def send_menu(message):
    bot.send_message(message.chat.id, '/photo Команда для составления фото отчета.\n'
    '/document Команда для конвертирования файлов(pdf <=> docx)\n'
    '/video Команда для конвертирования видеофайлов(MP4 <=> Webm)\n'
    '/audio Команда для конвертирования аудиофайлов(Wav <=> Ogg)')


@bot.message_handler(commands=['photo'])
def main(message):
    bot.send_message(message.chat.id, 'Отправьте этикетку для составления отчета')


@bot.message_handler(commands=['document'])
def main(message):
    bot.send_message(message.chat.id, 'Отправьте файл для конвертации')


@bot.message_handler(commands=['video'])
def main(message):
    bot.send_message(message.chat.id, 'Отправьте видеофайл для конвертации')


@bot.message_handler(commands=['audio'])
def main(message):
    bot.send_message(message.chat.id, 'Отправьте аудиофайл для конвертации')


@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    # Получаем информацию о фото
    file_info = bot.get_file(message.photo[-1].file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Сохраняем изображение во временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_image:
        temp_image.write(downloaded_file)
        temp_image_path = temp_image.name

    # Извлекаем текст с изображения
    result = reader.readtext(temp_image_path)

    # Удаляем временный файл
    os.remove(temp_image_path)

    # Создаем документ Word и записываем извлеченный текст
    doc = Document()
    for detection in result:
        text = detection[1]
        doc.add_paragraph(text)

    # Сохраняем документ Word во временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_doc:
        doc.save(temp_doc.name)
        doc_filename = temp_doc.name

    # Отправляем файл пользователю
    with open(doc_filename, "rb") as doc_file:
        bot.send_document(message.chat.id, doc_file)

    # Удаляем временный файл документа
    os.remove(doc_filename)


@bot.message_handler(content_types=['document'])
def handle_document(message):
    # Получаем информацию о документе
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Определяем формат документа
    file_extension = message.document.file_name.split('.')[-1]

    if file_extension == 'pdf':
        # Конвертируем PDF в DOCX
        converted_docx = convert_pdf_to_docx(downloaded_file)
        # Сохраняем конвертированный документ во временный файл
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_doc:
            temp_doc.write(converted_docx.getvalue())
            temp_docx_path = temp_doc.name
        # Отправляем конвертированный документ пользователю как файл с расширением .docx
        bot.send_document(message.chat.id, open(temp_docx_path, "rb"), caption="Конвертированный документ.docx")
        os.remove(temp_docx_path)

    elif file_extension == 'docx':
        # Сохраняем docx файл во временный файл
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_doc:
            temp_doc.write(downloaded_file)
            temp_docx_path = temp_doc.name

        # Конвертируем DOCX в PDF
        converted_pdf = convert_docx_to_pdf(temp_docx_path)
        # Отправляем конвертированный документ пользователю как файл с расширением .pdf
        bot.send_document(message.chat.id, open(converted_pdf, "rb"), caption="Конвертированный документ.pdf")
        os.remove(converted_pdf)
        os.remove(temp_docx_path)

    else:
        bot.reply_to(message, "Поддерживаются только документы в форматах PDF и DOCX")


@bot.message_handler(content_types=['audio'])
def handle_audio(message):
    # Получаем информацию об аудио сообщении
    file_info = bot.get_file(message.audio.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Определяем формат аудио
    file_extension = message.audio.mime_type.split('/')[-1]

    if file_extension == 'wav':
        # Конвертируем аудио из WAV в OGG
        converted_audio = convert_audio(downloaded_file, 'wav', 'ogg')
        bot.send_voice(message.chat.id, converted_audio)
    elif file_extension == 'ogg':
        # Конвертируем аудио из OGG в WAV
        converted_audio = convert_audio(downloaded_file, 'ogg', 'wav')
        bot.send_document(message.chat.id, converted_audio, caption="Конвертированный аудиофайл.wav")
    else:
        bot.reply_to(message, "Поддерживаются только аудиоформаты WAV и OGG")


@bot.message_handler(content_types=['video'])
def handle_video(message):
    # Получаем информацию о видео сообщении
    file_info = bot.get_file(message.video.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Определяем формат видео
    file_extension = message.video.mime_type.split('/')[-1]

    if file_extension == 'mp4':
        # Конвертируем видео из MP4 в WebM
        converted_video = convert_video(downloaded_file, 'mp4', 'webm')
        bot.send_document(message.chat.id, converted_video, caption="Конвертированный видеофайл.webm")
    elif file_extension == 'webm':
        # Конвертируем видео из WebM в MP4
        converted_video = convert_video(downloaded_file, 'webm', 'mp4')
        bot.send_document(message.chat.id, converted_video, caption="Конвертированный видеофайл.mp4")
    else:
        bot.reply_to(message, "Поддерживаются только видеоформаты MP4 и WebM")


def convert_audio(audio_bytes, from_format, to_format):
    audio = AudioSegment.from_file(io.BytesIO(audio_bytes), format=from_format)
    output_stream = io.BytesIO()
    audio.export(output_stream, format=to_format)
    output_stream.seek(0)
    return output_stream


def convert_video(video_bytes, from_format, to_format):
    input_stream = io.BytesIO(video_bytes)
    output_stream = io.BytesIO()

    with tempfile.NamedTemporaryFile(delete=False, suffix="." + from_format) as temp_input:
        temp_input.write(input_stream.read())
        temp_input_path = temp_input.name

    output_filename = "temp_output." + to_format

    (
        ffmpeg
        .input(temp_input_path)
        .output(output_filename)
        .run()
    )

    with open(output_filename, "rb") as f:
        output_stream.write(f.read())

    output_stream.seek(0)
    os.remove(temp_input_path)
    os.remove(output_filename)
    return output_stream


def convert_pdf_to_docx(pdf_bytes):
    pdf_reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))  # Замена PdfFileReader на PdfReader
    doc = Document()

    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        doc.add_paragraph(page.extract_text())

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream


def convert_docx_to_pdf(docx_path):
    # Конвертируем DOCX в PDF используя docx2pdf
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    return pdf_path


bot.polling(none_stop=True)