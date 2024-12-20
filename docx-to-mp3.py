import boto3
from docx import Document
import sys  # Importamos sys para poder terminar la ejecución con sys.exit()
import botocore.exceptions  # Añadimos la importación necesaria para manejar excepciones de botocore
import os


def sanitize_ssml_text(text):
    """
    Sanitizes the SSML (Speech Synthesis Markup Language) text in the given chunk to ensure it is properly formatted for use with the Amazon Polly text-to-speech service.

    This function replaces XML special characters with their text equivalents to avoid parsing errors.
    """
    # Replace XML special characters
    replacements = {
        '&': 'y',
        '<': ' menor que ',
        '>': ' mayor que ',
        '"': ' comillas ',
        "'": ' apostrofe ',
        # Add more replacements if needed
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    return text


def read_docx(file_path):
    """
    Reads the content of a Word document (.docx) file and returns it as a single string.

    Args:
        file_path (str): The path to the .docx file.
    """
    # Clean the file path but preserve intentional spaces in directory/file names
    cleaned_path = file_path.strip().strip("'").strip('"')
    # Handle path with spaces correctly
    cleaned_path = os.path.expanduser(os.path.expandvars(cleaned_path))

    if not os.path.exists(cleaned_path):
        raise FileNotFoundError(f"The file '{cleaned_path}' does not exist")

    doc = Document(cleaned_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    # Replace ampersands '&' with 'and' to avoid Polly exceptions
    full_text = [text.replace('&', 'and') for text in full_text]
    return ' '.join(full_text)


def choose_voice_and_rate():
    """
    Presents the user with a choice of voices and their corresponding speech rates,
    and returns the selected voice and rate.
    """
    voices = {
        'a': ('Lucia', '85%'),
        'b': ('Conchita', '99%'),
        'c': ('Enrique', '92%')
    }

    print("Please choose a voice:")
    for key, value in voices.items():
        print(f"{key}) {value[0]}")

    choice = input("Enter your choice (a, b, or c): ").lower()
    return voices.get(choice, ('Lucia', '85%'))  # Devuelve 'Lucia' y '85%' como valores predeterminados


def split_text(text, max_length):
    """
    Splits a given text into chunks of a maximum length, trying to split at spaces
    to avoid cutting words in half.

    Args:
        text (str): The text to split.
        max_length (int): The maximum length of each chunk.
    """
    chunks = []
    while text:
        if len(text) <= max_length:
            chunks.append(text)
            break
        else:
            # Buscamos el último espacio en blanco antes del límite de longitud
            split_index = text.rfind(' ', 0, max_length)
            if split_index == -1:  # No se encontró un espacio, cortamos en el límite máximo
                split_index = max_length
            chunks.append(text[:split_index])
            text = text[split_index + 1:]  # +1 para no incluir el espacio en el nuevo fragmento
    return chunks


def convert_docx_to_mp3():
    """
    Converts a Word document (.docx) to an MP3 audio file using Amazon Polly.
    It prompts the user for the file path, voice, and then performs the conversion.
    """
    file_path = input('Please enter the path to the Word document or type "exit" to abort: ')

    if file_path.lower() == 'exit':
        print('Program aborted by the user.')
        sys.exit()
    voice_id, rate = choose_voice_and_rate()

    text = read_docx(file_path)
    # Clean and normalize the output path
    cleaned_path = file_path.strip().strip("'").strip('"')
    cleaned_path = os.path.expanduser(os.path.expandvars(cleaned_path))
    mp3_output = cleaned_path.rsplit('.', 1)[0] + '.mp3'

    # Check if the output file exists and delete it if it does
    if os.path.exists(mp3_output):
        os.remove(mp3_output)

    polly_client = boto3.Session(region_name='eu-central-1').client('polly')
    max_chunk_len = 1000
    chunks = split_text(text, max_chunk_len)
    stream = bytearray()

    for i, chunk in enumerate(chunks, start=1):
        print(f"Chunk {i}: {chunk}")  # Imprime el fragmento de texto actual
        sanitized_chunk = sanitize_ssml_text(chunk)
        ssml_text = f'<speak><prosody rate="{rate}">{sanitized_chunk}</prosody></speak>'
        try:
            response = polly_client.synthesize_speech(VoiceId=voice_id,
                                                      OutputFormat='mp3',
                                                      TextType='ssml',
                                                      Text=ssml_text)
            stream += response['AudioStream'].read()
            sys.stdout.write(f'\rProcessing chunk {i}/{len(chunks)}...')
            sys.stdout.flush()
        except botocore.exceptions.ClientError as e:
            print(f"\nError en el chunk {i}: {e}")
            break  # Interrumpe el bucle si hay un error

    sys.stdout.write('\n')

    with open(mp3_output, 'wb') as f:
        f.write(stream)

    print(f'Conversion completed. MP3 output saved as {mp3_output}')


convert_docx_to_mp3()


