import pandas as pd
import os
import re

print("请确保处理的 Excel 文件中都存在 “存货编码“ , “色号” 和 “数量” 三列")
print("请输入需要分组的 Excel 文件的名称（包括扩展名，如 .xlsx），确保excel文件在当前目录下。")
required_columns1 = ['存货编码', '色号', '数量']
required_columns2 = ['存货编码', '色号']

while True:
    filename1 = input()
    file_name1, file_extension1 = os.path.splitext(filename1)
    try:
        if file_extension1==".xlsx":
            df = pd.read_excel(filename1, engine='openpyxl')
        else:
            df = pd.read_excel(filename1,engine='xlrd')
    except FileNotFoundError:
        print(f"文件 {filename1} 未找到，请确认文件名正确。")
        print("请重新输入:")
        continue
    except Exception as e:
        print(f"加载文件时发生错误: {e}")
        print("请重新输入:")
        continue
    missing_columns = [col for col in required_columns1 if col not in df.columns]
    if missing_columns:
        print(f"缺少以下列: {', '.join(missing_columns)}")
        print("请确保excel文件中存在以下列：",required_columns1)
        input("程序已结束，按任意键退出...")
        exit()
    else:
        print("所有必需的列都存在。")
        break
    
print("请输入用于查找存货编码的 Excel 文件的名称（包括扩展名，如 .xlsx），确保excel文件在当前目录下。")
while True:
    filename2 = input()
    file_name2, file_extension2 = os.path.splitext(filename2)
    try:
        if file_extension2==".xlsx":
            dfz = pd.read_excel(filename2, engine='openpyxl')
        else:
            dfz = pd.read_excel(filename2,engine='xlrd')
    except FileNotFoundError:
        print(f"文件 {filename2} 未找到，请确认文件名正确。")
        print("请重新输入:")
        continue
    except Exception as e:
        print(f"加载文件时发生错误: {e}")
        print("请重新输入:")
        continue
    missing_columns = [col for col in required_columns2 if col not in dfz.columns]
    if missing_columns:
        print(f"缺少以下列: {', '.join(missing_columns)}")
        print("请确保excel文件中存在以下列：",required_columns2)
        input("程序已结束，按任意键退出...")
        exit()
    else:
        print("所有必需的列都存在。")
        break

wlbh1 = df.columns.get_loc("存货编码")
wlbh2 = dfz.columns.get_loc("存货编码")
sh1=df.columns.get_loc("色号")
sh2=dfz.columns.get_loc("色号")
sl=df.columns.get_loc("数量")
error=[]
eind=[]

print("请输入想要复制或者添加的列名，若没有，请输入 0 ：")
col=input()
while(col!="0"):
    print("如果想要复制，请输入 1 ，若想要添加，请输入 2 ：")
    flag=int(input())
    if flag==1:
        if col not in dfz.columns:
            print(filename2+"中 "+col+" 列不存在，请重新输入！")
        else:
            if  col not in df.columns:
                print(filename1+"中 "+col+" 列不存在，请先添加 "+col+" 列再进行复制！")
            else:
                print("请输入你想要复制的值")
                opt=input()
                print("正在复制中，请稍候")
                df[col] = df[col].astype('object')
                for i in range(len(df)):
                    print("正在复制第 "+str(i+1)+" 行")
                    for j in range(len(dfz)):
                        if df.at[i, "色号"] == dfz.at[j, "色号"] and dfz.at[j, col] == opt:
                            df.at[i, col] = opt
                            break
                print("复制完成！")       
    else:
        if col in df.columns:
            print(filename1+"中 "+col+" 列已经存在，请重新输入！")
        else:
            print("请输入想要添加到第几列，如果此列已经存在数据，则原数据会后移,若想要添 加到最后一列，请输入 0 ：")
            ind=int(input())
            print("请输入该列需要插入的值（统一值）：")
            column_value = int(input())
            print("正在添加中，请稍候")
            if ind==0:
                df.insert(len(df.columns),col,column_value)
            else:
                df.insert(ind-1,col,column_value)
            print("添加完成！")
    print("请输入想要复制或者添加的列名，若没有，请输入 0 ：")
    col=input()

df1 = pd.DataFrame(columns=df.columns)
print("请输入想要多少为一组：")
z=int(input())
print('正在处理中，请稍候...')
cnt=1 

for i in df.values:
    print("正在处理第 "+str(cnt)+" 行")
    fg=0
    for j in dfz.values:
        if j[sh2]==i[sh1]:
            i[wlbh1]=j[wlbh2]
            fg=1
            break
    if fg==0:
        error.append(i[sh1])
        eind.append(cnt)
    cnt+=1
    if i[sl]>z:
        a=int(i[sl]/z)
        b=i[sl]%z
        i[sl]=z
        for j in range(0,a,1):
           df1.loc[len(df1)] = i
        i[sl]=b
        df1.loc[len(df1)] = i
    else:
        df1.loc[len(df1)] = i
df1.to_excel(file_name1+"output.xlsx",index=False)
print("处理完成！已生成"+file_name1+"output.xlsx"+"文件。")
print()
with open('output.txt', 'w', encoding='utf-8') as file:
    print("以下色号未能匹配到物料编号：", file=file)
    for i in range(0,len(error),1):
        print("第"+str(eind[i])+"行："+"色号为 "+str(error[i])+" 未能匹配到物料编号。",end="\n", file=file)
print("未能匹配到物料编号的色号已存储在同一目录下output.txt文件中。")
input("程序已结束，按任意键退出...")