import openai
import configparser
from urllib.parse import urlparse
import os
class GPTHandler:
    def __init__(self):
        config = configparser.ConfigParser()
        config.read("config.ini")
        api_key = config.get("api", "key")
        if api_key == "your_api_key":
            raise ValueError("Invalid API key. Please set a valid API key in the config.ini file.")
        proxy_url = config.get('DEFAULT', 'proxy')
        parsed_url = urlparse(proxy_url)
        if all([parsed_url.scheme, parsed_url.netloc]):
            self.proxy_url = proxy_url
        else:
            self.proxy_url = None
        openai.api_key = api_key

    def send_promt(self1):
        # 获取用户输入的文本
        user_input = self1.txt_input.get("1.0", "end-1c")
        # 生成 prompt
        excel_description=self1.excel_handler.generate_excel_description(self1.file_path)
        file_name, file_extension = os.path.splitext(self1.file_path)
        new_file_name = file_name + "处理后" + file_extension
        file_path=self1.file_path
        prompt = "该excle表的详细描述为"+excel_description+"路径为"+file_path+f"\n将以下需求转换为代码，代码要求：1.不能包含导入包的代码；2.不能包含除代码和注释意外的内容，3.代码处理思路是将excel表读到dataframe中在dataframe中处理保存在下面路径当中\n{new_file_name}：\n{user_input}"
        # 创建 GPTHandler 实例
        gpt_handler = GPTHandler()
        # 使用 gpt 来生成代码
        generated_code = gpt_handler.generate_code(prompt)
        # 将生成的代码插入到 txt_code
        self1.txt_code.insert("end", generated_code)
        # 生成解释的 prompt
        explain_prompt = f"解释以下代码,若该代码中含有读取文件的代码，记得提醒用户将路径名换成自己文件的路径名：\n{generated_code}"
        # 使用 gpt 来生成解释
        generated_explain = gpt_handler.generate_code(explain_prompt)
        # 将生成的解释插入到 txt_explain
        self1.txt_explain.insert("end", generated_explain)
    def generate_code(self, prompt):
        if self.proxy_url:
            proxies = {
                'http': self.proxy_url,
                'https': self.proxy_url
            }
        else:
            proxies = None

            # Use proxies in your API request if available
        response = openai.Completion.create(
            engine="gpt-3.5-turbo-instruct",
            prompt=prompt,
            max_tokens=500,
            n=1,
            stop=None,
            temperature=0.5
        )
        return response.choices[0].text.strip()