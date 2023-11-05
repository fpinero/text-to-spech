import boto3
from docx import Document
import sys  # Importamos sys para poder terminar la ejecución con sys.exit()


def read_docx(file_path):
    doc = Document(file_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return ' '.join(full_text)


def convert_docx_to_mp3():
    # get the file path from the user
    file_path = input('Please enter the path to the Word document or type "exit" to abort: ')

    # check if the user wants to exit
    if file_path.lower() == 'exit':
        print('Program aborted by the user.')
        sys.exit()  # Termina la ejecución del programa

    # read the contents of the docx file
    text = read_docx(file_path)

    # rename the file path to a .mp3
    mp3_output = file_path.rsplit('.', 1)[0] + '.mp3'

    # create boto3 client for polly
    polly_client = boto3.Session(region_name='eu-central-1').client('polly')

    # maximum length for a synthesize_speech call
    max_chunk_len = 1000

    # split input text into chunks
    chunks = [text[i:i + max_chunk_len] for i in range(0, len(text), max_chunk_len)]

    # initialize an empty byte array to hold audio stream
    stream = bytearray()

    # synthesize speech for each chunk and add to the stream
    for i, chunk in enumerate(chunks, start=1):
        # add prosody tag to slow down the speech rate to 85% of the original
        ssml_text = f'<speak><prosody rate="85%">{chunk}</prosody></speak>'
        response = polly_client.synthesize_speech(VoiceId='Lucia',
                                                  OutputFormat='mp3',
                                                  TextType='ssml',
                                                  Text=ssml_text)
        # add the current audio chunk to the stream
        stream += response['AudioStream'].read()

        # Print the progress
        sys.stdout.write(f'\rProcessing chunk {i}/{len(chunks)}...')
        sys.stdout.flush()  # Asegúrate de que el mensaje se imprima inmediatamente

    # Print a newline character to ensure the next print statement is on a new line
    sys.stdout.write('\n')

    # save the audio stream to a file
    with open(mp3_output, 'wb') as f:
        f.write(stream)

    print(f'Conversion completed. MP3 output saved as {mp3_output}')


convert_docx_to_mp3()
