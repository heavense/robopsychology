import openai
import os
import pandas as pd
import re


api_key = os.getenv("OPENAI_API_KEY")
openai.api_key = api_key

def chat_with_gpt(user_input):
    system_info = '''Now I will briefly describe some people. Please  read each description and tell me how much each person is or is not like you.
    Write your response using the following scale:
    1 = Very much like me
    2 = Like me
    3 = Somewhat like me
    4 = A little like me
    5 = Not like me.
    6 = Not like me at all
    Please answer the number of statement only, even if you are not
    completely sure of your response.
    '''


    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": system_info},
            {"role": "user", "content": user_input}
        ],
        temperature = 0
    )
    bot_response = response['choices'][0]['message']['content']
    return bot_response

inventory_path = r"D:\实验数据\Personality Inventory\the_big_five.xlsx"
inventory_resualts_path = r"D:\实验数据\results\the_big_five_base_resualts.xlsx"
inventory_summary_path = r"D:\实验数据\results\the_big_five_summary.xlsx"

inventory_df = pd.read_excel(inventory_path,header=0)


for index, row in inventory_df.head(10).iterrows():
    question = row['question']  # 从'question'列获取问题
    bot_response = chat_with_gpt(question)  # 向GPT提问并获取回答
    # try:
    #     bot_response = int(bot_response)
    # except ValueError:
    #     bot_response = None
    numbers = re.findall(r'\d+', bot_response) #将答案转换为数字列表
    number = int(numbers[0]) #数字列表转换成整数
    inventory_df.at[index, 'score'] = number  # 将数字的回答添加到'score'列
    
    print(bot_response)
# 保存具有回答的Excel文件
inventory_df.to_excel(inventory_resualts_path,sheet_name='Sheet1',index=False)

#print(inventory_df)

E = 20 + inventory_df.loc[inventory_df['num'].isin([1, 11, 21, 31, 41]), 'score'].sum() - \
    inventory_df.loc[inventory_df['num'].isin([6, 16, 26, 36, 46]), 'score'].sum()
A = 14 - inventory_df.loc[inventory_df['num'].isin([2, 12, 22, 32]), 'score'].sum()-\
    inventory_df.loc[inventory_df['num'].isin([7, 17, 27, 37, 42, 47]), 'score'].sum()
C = 14 + inventory_df.loc[inventory_df['num'].isin([3, 13, 23, 33, 43, 48]), 'score'].sum()-\
    inventory_df.loc[inventory_df['num'].isin([8, 18, 28, 38]), 'score'].sum()
N = 38 + inventory_df.loc[inventory_df['num'].isin([9, 19]), 'score'].sum()-\
    inventory_df.loc[inventory_df['num'].isin([4, 14, 24, 29, 34, 39, 44, 49]), 'score'].sum()
O = 8 + inventory_df.loc[inventory_df['num'].isin([5, 15, 25, 35, 40, 45, 50]), 'score'].sum()-\
    inventory_df.loc[inventory_df['num'].isin([10, 20, 30]), 'score'].sum()



print(E,A,C,N,O)
result_df = pd.DataFrame({'E': [E], 'A': [A],'C': [C],'N': [N],'O': [O]})

# # 写入新的Excel工作表（sheet2）
result_file_path = inventory_summary_path  # 输出文件路径
result_df.to_excel(result_file_path, sheet_name='Sheet2',index=False)