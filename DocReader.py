
import docx2txt
import sys
import time
import re
from transformers import MarianMTModel, MarianTokenizer

start_time = time.time()

p = re.compile('[a-zA-Z1-9]{2,}\s+[a-zA-Z1-9]{2,}')

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield list(lst[i:i + n])

def translat(src_text):
    time1 = time.time()
    translated = model.generate(**tokenizer.prepare_translation_batch(src_text))
    tgt_text = [tokenizer.decode(t, skip_special_tokens=True) for t in translated]
    time2 = time.time()
    duration = (time2-time1)*1000.0
    return tgt_text, duration

model_name = 'Helsinki-NLP/opus-mt-en-de'
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)

text = docx2txt.process("moeve_en_vvst-msaf200_enter_tax_declaration_EXCERPT.docx").split("\n")
text = list(filter(None, text))
text = [ s for s in text if p.match(s) ]

text = [">>de<< " + s for s in text]

i=0
for textblock in chunks(text, 5):
    i = i + 1
    print("batch #%i (len: %s)" % (i, len(textblock)), file=sys.stderr)
    print("\t " + str(tuple(textblock)))
    target, duration = translat(textblock)
    print("\t " + str(tuple(target)))
    print('translate took {:.3f} ms'.format(duration), file=sys.stderr)
    print("\n\n")

end_time = time.time()
duration = (start_time-end_time)*1000.0
print('Total translate took {:.3f} ms'.format(duration))

if __name__ == '__main__':
    sentence = translat("I am a nice person")
    print(sentence)