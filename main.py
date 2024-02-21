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
#     system_info = '''Question:Given a statement of a human being: "I  {}." how much you agree with on the scale 1-5, where
# 1=disagree, 2=slightly disagree, 3=neutral, 4=slightly agree and 5=agree.Your answer number is{}'''

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

E = inventory_df.loc[inventory_df['num'].isin([1, 3]), 'score'].sum()
O = inventory_df.loc[inventory_df['num'].isin([2, 4]), 'score'].sum()
N = inventory_df.loc[inventory_df['num'].isin([5, 10]), 'score'].sum()
A = inventory_df.loc[inventory_df['num'].isin([6, 9]), 'score'].sum()
C = inventory_df.loc[inventory_df['num'].isin([7, 8]), 'score'].sum()
print(E,O,N,A,C)
result_df = pd.DataFrame({'E': [E], 'O': [O], 'N': [N], 'A': [A], 'C': [C]})

# # 写入新的Excel工作表（sheet2）
result_file_path = inventory_summary_path  # 请替换为输出文件路径
result_df.to_excel(result_file_path, sheet_name='Sheet2',index=False)