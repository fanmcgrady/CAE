import os

from docx import Document


# 构造新闻
def extract_news(name):
    myDocument = Document(name)

    text_list = []

    style = "<style>img{padding:0 !important;}p .para {text-indent:2em;line-height:2em;}</style>"

    p_header = "<p class='text'>"
    p_last_header = "<p class='text' style='text-align:right'>"
    p_footer = "</p>"

    news_html = []
    news_html.append(style)

    # 提取文本
    for paragraph in myDocument.paragraphs:
        if paragraph.text != "":
            text_list.append(paragraph.text)

    last_index = len(text_list) - 1

    for index, item in enumerate(text_list):
        if index == last_index and len(item) < 10:
            news_html.append(p_last_header)
        else:
            news_html.append(p_header)

        news_html.append(item)
        news_html.append(p_footer)

    print("".join(news_html))


# 构造教学工作简报
def extract_report(number, date, title, news):
    header_text = []
    header_text.append("<p style=\"text-align:center;line-height:200%\">")
    header_text.append(
        "    <strong><span style=\"font-size:27px;line-height:200%;font-family: &#39;微软雅黑&#39;,sans-serif;color:red\">四川大学网络空间安全学院</span></strong>")
    header_text.append("</p>")
    header_text.append("<p style=\"text-align:center;line-height:200%\">")
    header_text.append(
        "    <strong><span style=\"font-size:48px;line-height:200%;font-family: 楷体;color:red\">教学工作检查简报</span></strong>")
    header_text.append("</p>")
    header_text.append("<p style=\"text-align:center;line-height:200%\">")
    header_text.append(
        "    <strong><span lang=\"EN-US\" style=\"font-size:19px;line-height:200%;font-family:&#39;times new roman&#39;,serif\">2020</span></strong><strong><span style=\"font-size:19px;line-height: 200%;font-family:宋体\">年第</span></strong><strong><span lang=\"EN-US\" style=\"font-size:19px;line-height:200%;font-family:&#39;times new roman&#39;,serif\">" + number + "</span></strong><strong><span style=\"font-size:19px;line-height: 200%;font-family:宋体\">期</span></strong>")
    header_text.append("</p>")
    header_text.append("<p style=\"text-indent:7px;font-size:19px;line-height: 200%;font-family:楷体\">")
    header_text.append(
        "    <strong>教学科编印</strong><strong><span lang=\"EN-US\" style=\"font-size:19px;line-height:200%;font-family:&#39;times new roman&#39;,serif\">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></strong><strong>" + date + "</strong>")
    header_text.append("</p>")
    header_text.append("<div style=\"line-height:200%;border-top:5px solid red;margin-bottom:10px;\">")
    header_text.append("    <strong> </strong>")
    header_text.append("</div>")
    header_text.append("<p style=\"text-align:center;line-height:29px;font-size:19px;\">")
    header_text.append("    <strong>" + title + "</strong>")
    header_text.append("</p>")
    header_text.append("<p style=\"text-align:center;line-height:29px\">")
    header_text.append("    <br/>")
    header_text.append("</p>")

    print("".join(header_text))
    extract_news(news)

if __name__ == '__main__':
    files = os.listdir("news")

    for index, file in enumerate(files):
        news = os.path.join("news", file)
        if os.path.isfile(news):
            print("处理文件：{}".format(news))
            extract_report("9", "2020年09月01日", "成都市中小学网络安全宣传周专题讲座", news)
