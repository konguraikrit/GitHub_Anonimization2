from pythainlp.tag import NER

tagger = NER("thainer")
text = "นายสมชาย ใจดี ทำงานที่การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย ในกรุงเทพมหานคร"
result = tagger.tag(text)
print("Raw tags:")
for word, tag in result:
    print(f"  {word!r:30} -> {tag}")
