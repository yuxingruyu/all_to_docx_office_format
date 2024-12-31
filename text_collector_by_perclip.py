# Create:20241231
# python编写一个或多个类，包含如下功能，
# 允许用户一次性输入（复制到）多行文字内容，
# 将这些文字保存到一个字符串中。

class TextCollector:
    def __init__(self):
        self.text = ""

    def input_text(self):
        print("请输入多行文字内容（结束输入请按Ctrl+D（Linux、Mac）或Ctrl+Z（Windows）后回车）：")
        lines = []
        while True:
            try:
                line = input()
                lines.append(line)
            except EOFError:
                break
        # 去除文本两端可能存在的空白字符，并按换行符分割成列表，再重新拼接成字符串，确保格式良好
        lines = [line.strip() for line in lines]
        self.text = "\n".join(lines)

    def get_text(self):
        return self.text

if __name__ == "__main__":

    collector = TextCollector()
    collector.input_text()
    print("你输入的内容如下：")
    print(collector.get_text())
