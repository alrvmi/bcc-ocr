from paddleocr import PaddleOCR

ocr = PaddleOCR(use_textline_orientation=True, lang='ru')

img_path = './data/scans/8A16.jpg'
result = ocr.predict(img_path)

# Достаём только текст и confidence
texts = result[0]['rec_texts']
scores = result[0]['rec_scores']

for text, score in zip(texts, scores):
    print(f"{text} (уверенность={score:.2f})")
