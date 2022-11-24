#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import networkx as nx
from addEdge import addEdge
import os
import jieba
import jieba.analyse
from collections import Counter
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
from scipy.stats import linregress
from IPython import embed
named_colorscales = px.colors.named_colorscales()


# In[2]:


score = pd.DataFrame({'公司':[],
                      '第10款':[],
                      '第31款':[],
                      '第51款':[],
                      '第11款':[],
                      '第14款':[],
                      '第20款':[],
                      '第23款':[],
                      '第8款':[],
                      '平均收帳天數':[],
                      '現金流動比率':[],
                      '每股盈餘':[],
                      '借款依存度':[],
                      '母公司背書保證佔淨值比':[],
                      '母公司資金貸放佔淨值比':[],
                      'M-Score':[],
                      '結構異常':[],
                      '董監持股':[],
                      '董監質押股':[]})


# In[3]:


score = pd.read_csv('score.csv')


# In[4]:


#### Read Data ####

#Indicator
df_raw = pd.read_excel("./data/指標數值.xlsx", engine='openpyxl', sheet_name = None)
xls = pd.ExcelFile("./data/指標數值.xlsx",engine='openpyxl')
sheet_names = xls.sheet_names

#Important Message
important_message = pd.read_csv('./data/important_massage.csv')
important_message = important_message.loc[:, ~important_message.columns.astype(str).str.contains('^Unnamed')]
important_message['條款'] = important_message['符合條款'].apply(lambda x: str(x.split(' ')[1]))

#Company Sheet
corporate = pd.read_excel('./data/發生時間及對照公司(改).xlsx', engine='openpyxl')
corporate = corporate.loc[:, ~corporate.columns.astype(str).str.contains('^Unnamed')]
corporate['公司名稱'][2] = '齊民'
corporate['公司名稱'][13] = '吉祥全'

#Node
company_node = pd.read_csv('./data/company_node3.csv')
company_node[company_node['職稱'] == '重整人']['職稱'] = '董事'
company_node[company_node['職稱'] == '臨時管理人']['職稱'] = '董事'
company_node[company_node['職稱'] == '重整監督人']['職稱'] = '董事'
company_node[company_node['職稱'] == '常務董事  (獨立董事)']['職稱'] = '獨立董事'
company_node = company_node[company_node['姓名'] != '公司已廢止  ，董事會已不存在  ，依公司法第83條規定，清算人之就任、解任等均應向法院聲報；依民法第42條第1項規定，清算中之公司，係屬法院監督範疇。']
company_node = company_node[company_node['姓名'] != '董事']
company_node = company_node[company_node['姓名'] != '暫缺']
company_node = company_node.dropna(subset=['姓名']).reset_index(drop=True)
color_map = {'董事':'#3366CC', '獨立董事':'#FF9900', '代表人':'#4ecc63', '董事長':'#0099c6', '監察人':'#DD4477', '副董事長':'#316395', '公司':'#b07b4a'}
company_node['顏色'] = company_node['職稱']
company_node['顏色'] = company_node['顏色'].map(color_map)


# In[5]:


terms_of_corporate_governance = ['8']
terms_of_corporate_finance = ['14', '20', '11', '23']
terms_of_corporate_operating = ['31', '10', '51']


# In[6]:


t = [8,14, 20, 11, 23,31, 10, 51]


# In[7]:

for corporate_name, compare_name in zip(corporate['公司名稱'][14:], corporate['公司名稱'][:14]):
    #embed()
    # corporate_name = '歌林'
    # compare_name = '聲寶'


    # ### 資料準備

    # In[8]:


    #### Data Preparing ####

    score_dict = {}
    score_dict['公司'] = corporate_name
    df_corporate = df_raw.get(corporate_name)
    df_corporate = df_corporate.loc[:, ~df_corporate.columns.astype(str).str.contains('^Unnamed')].iloc[11:21, 2:7]
    df_corporate.columns = [str(d.year) for d in list(df_corporate)]

    important_message_corporate = important_message[important_message['公司'] == int(corporate[corporate['公司名稱'] == corporate_name]['股票代碼'].values[0])]
    important_message_corporate['年份'] = important_message_corporate['年份'].astype(str)



    df_compare = df_raw.get(compare_name)
    df_compare = df_compare.loc[:, ~df_compare.columns.astype(str).str.contains('^Unnamed')].iloc[11:21, 2:7]
    df_compare.columns = [str(d.year) for d in list(df_compare)]

    important_message_compare = important_message[important_message['公司'] == int(corporate[corporate['公司名稱'] == compare_name]['股票代碼'].values[0])]
    important_message_compare['年份'] = important_message_compare['年份'].astype(str)


    # ### Corporate Operating

    # In[9]:


    #terms_of_corporate_operating
    important_message_operating_unusual = important_message_corporate[important_message_corporate ['條款'].isin(terms_of_corporate_operating)]
    important_message_operating_usual = important_message_corporate[~important_message_corporate.isin(important_message_operating_unusual)].dropna()
    important_message_operating_unusual.iloc[:,:10].to_csv('./visual_output/operating/original_table/{}_important_message_operating.csv'.format(corporate_name), encoding='utf-8-sig', header=None, index=False)

    score_dict['第10款'] = sum(important_message_operating_unusual['符合條款'].str.contains('10'))
    score_dict['第31款'] = sum(important_message_operating_unusual['符合條款'].str.contains('31'))
    score_dict['第51款'] = sum(important_message_operating_unusual['符合條款'].str.contains('51'))

    usual = important_message_operating_usual.groupby('年份').count().reset_index()
    unusual = important_message_operating_unusual.groupby('年份').count().reset_index()

    if len(usual) == 0:
        usual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

    usual['條款狀態'] = '正常'
    unusual['條款狀態'] = '有舞弊風險'

    group_df = usual.append(unusual)
    group_df.rename(columns={'公司':'重大訊息數量'}, inplace=True)

    fig = px.bar(group_df, x="年份", y="重大訊息數量", color="條款狀態", title="重大訊息-條款年份統計")
    fig.write_html("./visual_output/operating/important_message/{}_important_message_operating.html".format(corporate_name))
    #fig.show()


    # In[10]:


    # Statistic of unusual terms
    unusual = important_message_operating_unusual.groupby(['年份','條款']).count().reset_index().rename(columns={'公司':'舞弊訊息數量','條款':'第幾款'})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '舞弊訊息數量':[0,0,0,0,0], '條款':[0,0,0,0,0], '第幾款':[0,0,0,0,0]})

    fig = px.bar(unusual, x="年份", y="舞弊訊息數量", color="第幾款", title="重大訊息-潛在舞弊條款統計", color_discrete_sequence=px.colors.qualitative.G10[1:4])
    fig.write_html("./visual_output/operating/unusual_message/{}_unusual_message_operating.html".format(corporate_name))
    #fig.show()


    # ## Corporate Finace

    # In[11]:

    a = list(df_corporate.loc[11])[-4:]
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['平均收帳天數'] = 10
    else:
        score_dict['平均收帳天數'] = 0

    for i in a:
        if i<=50:
            continue
        elif 50<i<100:
            score_dict['平均收帳天數'] += 2
        elif i>=100:
            score_dict['平均收帳天數'] += 5

    # In[12]:


    #Average collection days
    title = ['本公司平均收帳天數', '對照公司平均收帳天數']
    labels = ['本公司平均', '整體平均']
    colors = ['rgb(49,130,189)', 'rgb(115,115,115)']

    mode_size = [12, 8]
    line_size = [4, 2]

    x_data = np.array([list(df_corporate)[-4:],list(df_compare)[-4:]])
    y_data = np.array([list(df_corporate.loc[11])[-4:],list(df_compare.loc[11])[-4:]])

    fig = go.Figure()

    for i in range(0, 2):
        fig.add_trace(go.Scatter(x=x_data[i], y=y_data[i], mode='lines',
            name=labels[i],
            line=dict(color=colors[i], width=line_size[i]),
            connectgaps=True,
        ))

        # endpoints
        fig.add_trace(go.Scatter(
            x=[x_data[i][0], x_data[i][-1]],
            y=[y_data[i][0], y_data[i][-1]],
            mode='markers',
            marker=dict(color=colors[i], size=mode_size[i])
        ))

    fig.update_layout(
        xaxis=dict(
            showline=True,
            showgrid=False,
            showticklabels=True,
            linecolor='rgb(204, 204, 204)',
            linewidth=2,
            ticks='outside',
            tickfont=dict(
                family='Arial',
                size=12,
                color='rgb(82, 82, 82)',
            ),
        ),
        yaxis=dict(
            showgrid=False,
            zeroline=False,
            showline=False,
            showticklabels=False,
        ),
        autosize=False,
        margin=dict(
            autoexpand=False,
            l=100,
            r=20,
            t=110,
        ),
        showlegend=False,
        plot_bgcolor='white'
    )

    annotations = []

    # Adding labels

    # labeling the left_side of the plot
    annotations.append(dict(xref='paper', x=0.05, y=y_data[0][0],
                                  xanchor='right', yanchor='top',
                                  text=labels[0] + ' {}'.format(y_data[0][0]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the right_side of the plot
    annotations.append(dict(xref='paper', x=0.95, y=y_data[0][3],
                                  xanchor='left', yanchor='middle',
                                  text='{}'.format(y_data[0][3]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the left_side of the plot
    annotations.append(dict(xref='paper', x=0.05, y=y_data[1][0],
                                  xanchor='right', yanchor='bottom',
                                  text=labels[1] + ' {}'.format(y_data[1][0]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the right_side of the plot
    annotations.append(dict(xref='paper', x=0.95, y=y_data[1][3],
                                  xanchor='left', yanchor='middle',
                                  text='{}'.format(y_data[1][3]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))

    # Title
    annotations.append(dict(xref='paper', yref='paper', x=0.5, y=1.05,
                                  xanchor='center', yanchor='bottom',
                                  text='平均收帳天數',
                                  font=dict(family='Arial',
                                            size=16),
                                  showarrow=False))

    fig.update_layout(annotations=annotations)
    fig.write_html("./visual_output/finance/average_collection_days/{}_average_collection_days.html".format(corporate_name))
    #fig.show()


    # In[13]:


    a = list(df_corporate.loc[12])[-4:]
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope<0:
        score_dict['現金流動比率'] = 10
    else:
        score_dict['現金流動比率'] = 0
        
    for i in a:
        if i>100:
            continue
        elif 0<=i<=100:
            score_dict['現金流動比率'] += 2
        elif i<0:
            score_dict['現金流動比率'] += 5


    # In[14]:


    #Cash flow ratio
    title = ['本公司現金流動比率', '對照公司現金流動比率']
    labels = ['本公司比率', '整體平均比率']
    colors = ['rgb(49,130,189)', 'rgb(115,115,115)']

    mode_size = [8, 8]
    line_size = [4, 2]

    x_data = np.array([list(df_corporate)[-4:],list(df_compare)[-4:]])
    y_data = np.array([list(df_corporate.loc[12])[-4:],list(df_compare.loc[12])[-4:]])

    fig = go.Figure()

    for i in range(0, 2):
        x_red = []
        y_red = [d for d in y_data[i] if d<0]
        for index in y_red:
            x_red.append(np.where(y_data[i] == index)[0][0])

        x_yellow = []
        y_yellow = [d for d in y_data[i] if 0<=d<100]
        for index in y_yellow:
            x_yellow.append(np.where(y_data[i] == index)[0][0])
            
        fig.add_trace(go.Scatter(x=x_data[i], y=y_data[i], mode='lines',
            name=labels[i],
            line=dict(color=colors[i], width=line_size[i]),
            connectgaps=True,
        ))
        # yellow point
        fig.add_trace(go.Scatter(
            x=[v for i,v in enumerate(x_data[i].tolist()) if i in x_yellow],
            y=y_yellow,
            mode='markers',
            marker=dict(color='#ffb751', size=mode_size[i])
        ))
        
        # redpoint
        fig.add_trace(go.Scatter(
            x=[v for i,v in enumerate(x_data[i].tolist()) if i in x_red],
            y=y_red,
            mode='markers',
            marker=dict(color='#c52828', size=mode_size[i])
        ))

    fig.update_layout(
        xaxis=dict(
            showline=True,
            showgrid=True,
            showticklabels=True,
            linecolor='rgb(204, 204, 204)',
            linewidth=2,
            ticks='outside',
            tickfont=dict(
                family='Arial',
                size=12,
                color='rgb(82, 82, 82)',
            ),
        ),
        yaxis=dict(
            showgrid=False,
            zeroline=False,
            showline=False,
            showticklabels=False,
        ),
        autosize=False,
        margin=dict(
            autoexpand=False,
            l=100,
            r=20,
            t=110,
        ),
        showlegend=False,
        plot_bgcolor='white'
    )

    annotations = []

    # Adding labels

    # labeling the left_side of the plot
    annotations.append(dict(xref='paper', x=0.05, y=y_data[0][0],
                                  xanchor='right', yanchor='top',
                                  text=labels[0] + ' {}'.format(y_data[0][0]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the right_side of the plot
    annotations.append(dict(xref='paper', x=0.95, y=y_data[0][3],
                                  xanchor='left', yanchor='middle',
                                  text='{}'.format(y_data[0][3]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the left_side of the plot
    annotations.append(dict(xref='paper', x=0.05, y=y_data[1][0],
                                  xanchor='right', yanchor='bottom',
                                  text=labels[1] + ' {}'.format(y_data[1][0]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))
    # labeling the right_side of the plot
    annotations.append(dict(xref='paper', x=0.95, y=y_data[1][3],
                                  xanchor='left', yanchor='middle',
                                  text='{}'.format(y_data[1][3]),
                                  font=dict(family='Arial',
                                            size=13),
                                  showarrow=False))

    # Title
    annotations.append(dict(xref='paper', yref='paper', x=0.5, y=1.05,
                                  xanchor='center', yanchor='bottom',
                                  text='現金流動比率',
                                  font=dict(family='Arial',
                                            size=16,),
                                  showarrow=False))

    fig.update_layout(annotations=annotations)
    fig.write_html("./visual_output/finance/cash_flow_ratio/{}_cash_flow_ratio.html".format(corporate_name))
    #fig.show()


    # In[15]:


    #Earnings per share

    earnings_per_share = pd.DataFrame(df_corporate.loc[13]).reset_index().rename(columns={'index':'year', 13:'earnings per share'})

    colors_mark_line = ['lightslategray',]*len(earnings_per_share)
    for i,v in enumerate(earnings_per_share['earnings per share']):
        if v<0: colors_mark_line[i] = '#c52828'
        elif 0<v<5 : colors_mark_line[i] = '#ffb751'
        else: continue

    fig = go.Figure(data=[go.Bar(
        x=earnings_per_share['year'],
        y=earnings_per_share['earnings per share'],
        marker_color='lightslategray', # marker color can be a single color value or an iterable
        marker_line_color=colors_mark_line,
        marker_line_width=3,
    )])
    fig.update_layout(
        title_text='每股盈餘',
        xaxis=dict(
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        yaxis=dict(
            title='盈餘',
            titlefont_size=16,
            tickfont_size=14,
        ),)
    fig.write_html("./visual_output/finance/earnings_per_share/{}_earnings_per_share.html".format(corporate_name))
    #fig.show()


    # In[16]:


    a = list(earnings_per_share['earnings per share'])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope<0:
        score_dict['每股盈餘'] = 10
    else:
        score_dict['每股盈餘'] = 0
        
    for i in a:
        if i>5:
            continue
        elif 0<=i<=5:
            score_dict['每股盈餘'] += 2
        elif i<0:
            score_dict['每股盈餘'] += 5


    # In[17]:


    a = list(df_corporate.loc[17])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['母公司背書保證佔淨值比'] = 10
    else:
        score_dict['母公司背書保證佔淨值比'] = 0
        
    for i in a:
        if 0<=i<10:
            continue
        elif 10<=i<20:
            score_dict['母公司背書保證佔淨值比'] += 2
        elif (i>=20 or i<0):
            score_dict['母公司背書保證佔淨值比'] += 5
            
    a = list(df_corporate.loc[18])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['母公司資金貸放佔淨值比'] = 10
    else:
        score_dict['母公司資金貸放佔淨值比'] = 0
        
    for i in a:
        if 0<=i<10:
            continue
        elif 10<=i<20:
            score_dict['母公司資金貸放佔淨值比'] += 2
        elif (i>=20 or i<0):
            score_dict['母公司資金貸放佔淨值比'] += 5


    # In[18]:


    #parent company

    legend_x = 0
    legend_y = 1

    if (df_corporate.loc[17][0]>70 or df_corporate.loc[18][0]>70):
        legend_x = 0.78
        legend_y = 1

    if (df_corporate.loc[17][-1]>70 or df_corporate.loc[18][-1]>70):
        legend_x = 1
        legend_y = 1

    years = list(df_corporate)

    color_endorse = ['rgb(55, 83, 109)']*len(years)
    for i,v in enumerate(list(df_corporate.loc[17])):
        if (v>=20 or v<0): color_endorse[i] = '#c52828'
        elif 10<=v<20 : color_endorse[i] = '#ffb751'
        else: continue

    color_funds = ['lightslategray']*len(years)
    for i,v in enumerate(list(df_corporate.loc[18])):
        if (v>=20 or v<0): color_funds[i] = '#c52828'
        elif 10<=v<20 : color_funds[i] = '#ffb751'
        else: continue

    fig = go.Figure()

    fig.add_trace(go.Bar(x=years,
                    y=df_corporate.loc[17],
                    name='母公司背書保証佔淨值比',
                    marker_color='rgb(55, 83, 109)',
                    marker_line_color=color_endorse,
                    marker_line_width=3
                    ))
    fig.add_trace(go.Bar(x=years,
                    y=df_corporate.loc[18],
                    name='母公司資金貸放佔淨值比',
                    marker_color='lightslategray',
                    marker_line_color=color_funds,
                    marker_line_width=3,
                    ))

    fig.update_layout(
        title='母公司背書保証/資金貸放佔淨值比',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='佔比(%)',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        legend=dict(
            x=legend_x,
            y=legend_y,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        barmode='group',
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.write_html("./visual_output/finance/endorse_funds/{}_endorse_funds.html".format(corporate_name))
    #fig.show()


    # In[19]:


    #Borrowing dependence
    borrowing_dependence = pd.DataFrame(df_corporate.loc[16]).reset_index().rename(columns={16:'借款依存度', 'index':'年份'})

    color = ['#096148']*len(borrowing_dependence['年份'])
    for i,v in enumerate(list(borrowing_dependence['借款依存度'])):
        if (v<0 or v>200): color[i] = '#c52828'
        elif 100<=v<=200 : color[i] = '#ffb751'
        else: continue
            
    fig = go.Figure()

    fig.add_trace(dict(
        x=borrowing_dependence["年份"],
        y=borrowing_dependence["借款依存度"],
        hoverinfo='x+y',
        mode='lines',
        name = '借款依存度',
        line=dict(width=0.5,
                  color='#096148'),
        stackgroup='one'
    ))

    fig.add_trace(go.Scatter(
            x=borrowing_dependence["年份"],
            y=borrowing_dependence["借款依存度"],
            mode='markers',
            marker=dict(color=color, size=5)
        ))

    fig.update_layout(
        title='借款依存度',
        showlegend=False,
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='借款依存度',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        legend=dict(
            x=legend_x,
            y=legend_y,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        barmode='group',
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.write_html("./visual_output/finance/borrowing_dependence/{}_borrowing_dependence.html".format(corporate_name))
    #fig.show()


    # In[20]:


    a = list(df_corporate.loc[16])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['借款依存度'] = 10
    else:
        score_dict['借款依存度'] = 0
        
    for i in a:
        if 0<i<100:
            continue
        elif 100<=i<=200:
            score_dict['借款依存度'] += 2
        elif i<0 or i>200:
            score_dict['借款依存度'] += 5


    # In[21]:


    a = list(df_corporate.loc[19])[-2:]
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['M-Score'] = 4
    else:
        score_dict['M-Score'] = 0
        
    for i in a:
        if i<(-2.277):
            continue
        elif -2.277<=i<=(-1.837):
            score_dict['M-Score'] += 2
        elif i>(-1.837):
            score_dict['M-Score'] += 5


    # In[22]:


    # M score
    M_score = pd.DataFrame(df_corporate.loc[19]).reset_index().rename(columns={19:'M score', 'index':'年份'})

    years = list(M_score['年份'])[-2:]

    color = ['rgb(55, 83, 109)']*len(years)

    for i,v in enumerate(list(df_corporate.loc[19])[-2:]):
        if v>-1.837: color[i] = '#c52828'
        elif -2.277<=v<=-1.837 : color[i] = '#ffb751'
        else: continue
            
    fig = go.Figure()

    fig.add_trace(go.Bar(x=M_score['年份'][-2:],
                    y=M_score['M score'][-2:],
                    name='M score',
                    marker_color='rgb(55, 83, 109)',
                    marker_line_color=color,
                    marker_line_width=3
                    ))

    fig.update_layout(
        title='M score',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='M score',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        legend=dict(
            x=0,
            y=1,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        barmode='group',
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.write_html("./visual_output/finance/M_score/{}_M_score.html".format(corporate_name))
    #fig.show()


    # In[23]:


    #terms_of_corporate_finance
    important_message_finance_unusual = important_message_corporate[important_message_corporate ['條款'].isin(terms_of_corporate_finance)]
    important_message_finance_usual = important_message_corporate[~important_message_corporate.isin(important_message_finance_unusual)].dropna()
    important_message_finance_unusual.iloc[:,:10].to_csv('./visual_output/finance/original_table/{}_important_message_finance.csv'.format(corporate_name), encoding='utf-8-sig', header=None, index=False)

    score_dict['第11款'] = sum(important_message_finance_unusual['符合條款'].str.contains('11'))
    score_dict['第14款'] = sum(important_message_finance_unusual['符合條款'].str.contains('14'))
    score_dict['第20款'] = sum(important_message_finance_unusual['符合條款'].str.contains('20'))
    score_dict['第23款'] = sum(important_message_finance_unusual['符合條款'].str.contains('23'))

    usual = important_message_finance_usual.groupby('年份').count().reset_index()
    unusual = important_message_finance_unusual.groupby('年份').count().reset_index()

    if len(usual) == 0:
        usual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

            
    usual['條款狀態'] = '正常'
    unusual['條款狀態'] = '有舞弊風險'

    group_df = usual.append(unusual)
    group_df.rename(columns={'公司':'重大訊息數量'}, inplace=True)

    fig = px.bar(group_df, x="年份", y="重大訊息數量", color="條款狀態", title="重大訊息-條款年份統計")
    fig.write_html("./visual_output/finance/important_message/{}_important_message_finance.html".format(corporate_name))
    #fig.show()


    # In[24]:


    # Statistic of unusual terms
    unusual = important_message_finance_unusual.groupby(['年份','條款']).count().reset_index().rename(columns={'公司':'舞弊訊息數量','條款':'第幾款'})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '第幾款':[0,0,0,0,0]})

    x = sorted(list(unusual['年份'].unique()))

    fig = go.Figure()

    color = ['#DC3912', '#FF9900', 'indianred', 'lightsalmon']
    for i,t in enumerate(terms_of_corporate_finance):
        y_data = []
        unusual[unusual['第幾款'] == t]
        for y in x:
            try:
                y_data.append(int(unusual[unusual['第幾款'] == t][unusual[unusual['第幾款'] == t]['年份'] == y]['舞弊訊息數量']))
            except:
                y_data.append(0)      
        fig.add_trace(go.Bar(x=x, y=y_data, name=t, marker_color=color[i]))

    fig.update_layout(
        barmode='stack', 
        title='重大訊息-潛在舞弊條款統計',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='舞弊訊息數量',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            categoryorder='category ascending',
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        legend=dict(
            title='第幾款',
            x=1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.write_html("./visual_output/finance/unusual_message/{}_unusual_message_finance.html".format(corporate_name))
    #fig.show()


    # ## Governance

    # In[25]:


    #terms_of_corporate_governance
    important_message_governance_unusual = important_message_corporate[important_message_corporate ['條款'].isin(terms_of_corporate_governance)]
    important_message_governance_usual = important_message_corporate[~important_message_corporate.isin(important_message_governance_unusual)].dropna()
    important_message_governance_unusual.iloc[:,:10].to_csv('./visual_output/governance/original_table/{}_important_message_governance.csv'.format(corporate_name), encoding='utf-8-sig', header=None, index=False)

    score_dict['第8款'] = sum(important_message_governance_unusual['符合條款'].str.contains('8'))

    usual = important_message_finance_usual.groupby('年份').count().reset_index()
    unusual = important_message_finance_unusual.groupby('年份').count().reset_index()

    if len(usual) == 0:
        usual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '條款':[0,0,0,0,0]})

            
    usual['條款狀態'] = '正常'
    unusual['條款狀態'] = '有舞弊風險'

    group_df = usual.append(unusual)
    group_df.rename(columns={'公司':'重大訊息數量'}, inplace=True)

    fig = px.bar(group_df, x="年份", y="重大訊息數量", color="條款狀態", title="重大訊息-條款年份統計")
    fig.write_html("./visual_output/governance/important_message/{}_important_message_governance.html".format(corporate_name))
    #fig.show()


    # In[26]:


    # Statistic of unusual terms
    unusual = important_message_governance_unusual.groupby(['年份','條款']).count().reset_index().rename(columns={'公司':'舞弊訊息數量','條款':'第幾款'})

    if len(unusual) == 0:
        unusual = pd.DataFrame({'年份':[2005,2006,2007,2008,2009], '公司':[0,0,0,0,0], '第幾款':[0,0,0,0,0]})

            
    x = sorted(list(unusual['年份'].unique()))

    fig = go.Figure()

    color = ['indianred', 'lightsalmon', '#DC3912', '#FF9900']
    for i,t in enumerate(terms_of_corporate_governance):
        y_data = []
        unusual[unusual['第幾款'] == t]
        for y in x:
            try:
                y_data.append(int(unusual[unusual['第幾款'] == t][unusual[unusual['第幾款'] == t]['年份'] == y]['舞弊訊息數量']))
            except:
                y_data.append(0)      
        fig.add_trace(go.Bar(x=x, y=y_data, name=t, marker_color=color[i], showlegend=True))

    fig.update_layout(
        barmode='stack', 
        title='重大訊息-潛在舞弊條款統計',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='舞弊訊息數量',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            categoryorder='category ascending',
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        legend=dict(
            title='第幾款',
            x=1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)',
        ),
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )

    fig.write_html("./visual_output/governance/unusual_message/{}_unusual_message_operating.html".format(corporate_name))
    #fig.show()


    # In[27]:


    a = list(df_corporate.loc[14])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope<0:
        score_dict['董監持股'] = 10
    else:
        score_dict['董監持股'] = 0
        
    for i in a:
        if i>20:
            continue
        elif 10<i<20:
            score_dict['董監持股'] += 2
        elif i<=10:
            score_dict['董監持股'] += 5
            
    a = list(df_corporate.loc[15])
    b = list(range(len(a)))
    model = linregress(a, b) 
    slope, intercept = model.slope, model.intercept 

    if slope>0:
        score_dict['董監質押股'] = 10
    else:
        score_dict['董監質押股'] = 0
        
    for i in a:
        if i<=10:
            continue
        elif 10<i<33:
            score_dict['董監質押股'] += 2
        elif i>=33:
            score_dict['董監質押股'] += 5


    # In[28]:


    # Directors and supervisors

    title = '董監持股/質押股'
    labels = ['董監持股', '質押股']
    colors = ['rgb(67,67,67)', 'rgb(49,130,189)']

    mode_size = [10, 10]
    line_size = [3, 3]

    x_data = np.array([list(df_corporate), list(df_corporate)])

    y_data = df_corporate.loc[14:15].values

    fig = go.Figure()

    for i in range(0, 2):
        fig.add_trace(go.Scatter(x=x_data[i], y=y_data[i], mode='lines',
            name=labels[i],
            line=dict(color=colors[i], width=line_size[i]),
            connectgaps=True,
        ))

        # endpoints
        fig.add_trace(go.Scatter(
            x=[x_data[i][0], x_data[i][-1]],
            y=[y_data[i][0], y_data[i][-1]],
            mode='markers',
            marker=dict(color=colors[i], size=mode_size[i])
        ))

    fig.update_layout(
        xaxis=dict(
            showline=True,
            showgrid=False,
            showticklabels=True,
            linecolor='rgb(204, 204, 204)',
            linewidth=2,
            ticks='outside',
            tickfont=dict(
                family='Arial',
                size=12,
                color='rgb(82, 82, 82)',
            ),
        ),
        yaxis=dict(
            showgrid=False,
            zeroline=False,
            showline=False,
            showticklabels=False,
        ),
        autosize=False,
        margin=dict(
            autoexpand=False,
            l=100,
            r=20,
            t=110,
        ),
        showlegend=False,
        plot_bgcolor='white'
    )

    annotations = []

    # Adding labels
    for y_trace, label, color in zip(y_data, labels, colors):
        # labeling the left_side of the plot
        annotations.append(dict(xref='paper', x=0.05, y=y_trace[0],
                                      xanchor='right', yanchor='middle',
                                      text=label + ' {}'.format(y_trace[0]),
                                      font=dict(family='Arial',
                                                size=16),
                                      showarrow=False))
        # labeling the right_side of the plot
        annotations.append(dict(xref='paper', x=0.95, y=y_trace[4],
                                      xanchor='left', yanchor='middle',
                                      text='{}'.format(y_trace[4]),
                                      font=dict(family='Arial',
                                                size=16),
                                      showarrow=False))
    # Title
    annotations.append(dict(xref='paper', yref='paper', x=0.5, y=1.05,
                                  xanchor='center', yanchor='bottom',
                                  text='董監持股/質押股比例(%)',
                                  font=dict(family='Arial',
                                            size=16),
                                  showarrow=False))

    fig.update_layout(annotations=annotations)
    fig.write_html("./visual_output/governance/directors_supervisors/{}_directors_supervisors.html".format(corporate_name))
    #fig.show()


    # In[29]:


    # Network graph

    company_node_subset = company_node[company_node['公司'] == corporate[corporate['公司名稱'] == corporate_name]['公司全稱'].values[0]]

    company_list1 = list(set([i for i in list(set(company_node_subset['所代表法人'])) if type(i)!=float]))
    company_list2 = list(set([i for i in list(set(company_node_subset['公司'])) if type(i)!=float]))

    l = list(set(company_node_subset['姓名']))
    l.extend(company_list1)
    l.extend(company_list2)

    company_node_subset_map={}
    for i,v in enumerate(l):
        if v not in company_node_subset_map:
            company_node_subset_map[v] = i
    reverse_company_node_subset_map={}
    for k,v in company_node_subset_map.items():
        reverse_company_node_subset_map[v]=k

    company_node_subset['姓名'] = company_node_subset['姓名'].map(company_node_subset_map)
    company_node_subset['所代表法人'] = company_node_subset['所代表法人'].map(company_node_subset_map)
    company_node_subset['公司'] = company_node_subset['公司'].map(company_node_subset_map)
    company_node_subset['主要公司'] = company_node_subset['主要公司'].map(company_node_subset_map)

    company_list1.extend(company_list2)
    c_list = [company_node_subset_map[i] for i in company_list1]


    # In[30]:


    nodes = company_node_subset_map.values()
    edges = []
    for i,v in company_node_subset.iterrows():
        if str(v['主要公司']) != 'nan':
            edges.append((int(v['主要公司']), v['姓名']))
        if str(v['所代表法人']) != 'nan':
            edges.append((v['姓名'],int(v['所代表法人'])))

    #Target risk people
    target={}

    for e in list(set(edges)):
        if e[0] in c_list:
            if e[1] in target:
                target[e[1]] += 2
            else:
                target[e[1]] = 1

    target_people = []
    if len(c_list) >=3:
        for k,v in target.items():
            if v >= 0.75*len(c_list):
                target_people.append(k)
    
    if len(target_people)>=1:
        score_dict['結構異常'] = 1
    else:
        score_dict['結構異常'] = 0

    # Build Graph
    G = nx.Graph()
    for e in edges:
        G.add_edge(e[0],e[1])

    pos = nx.kamada_kawai_layout(G)

    fig = go.Figure()

    # Arrow Edge
    nodeColor = 'Blue'
    nodeSize = 20
    lineWidth = 3
    lineColor = '#000000'

    edge_x = []
    edge_y = []
    for edge in edges:
        start = pos[edge[0]]
        end = pos[edge[1]]
        edge_x, edge_y = addEdge(start, end, edge_x, edge_y, .8, 'end', .02, 30, nodeSize)

    edge_trace = go.Scatter(
        name="關係",
        x=edge_x, y=edge_y,
        line=dict(width=1, color='#888'),
        hoverinfo='none',
        mode='lines')

    fig.add_trace(edge_trace)

    # Node
    node_x = []
    node_y = []
    for node in G.nodes():
        x, y = pos[node][0], pos[node][1]
        node_x.append(x)
        node_y.append(y)

    director_x = []
    independent_director_x = []
    representative_x = []
    chairman_x = []
    supervisor_x = []
    vice_chairman_x = []
    corporate_x = []
    nothing_x = []

    director_y = []
    independent_director_y = []
    representative_y = []
    chairman_y = []
    supervisor_y = []
    vice_chairman_y = []
    corporate_y = []
    nothing_y = []

    director_index = []
    independent_director_index = []
    representative_index = []
    chairman_index = []
    supervisor_index = []
    vice_chairman_index = []
    corporate_index = []
    nothing_index = []

    for i, n in enumerate(list(G.nodes())):
        if n in list(company_node_subset['姓名']):
            title = company_node_subset[company_node_subset['姓名'] == n].values[0][2]
            if title == '董事長': 
                chairman_x.append(node_x[i])
                chairman_y.append(node_y[i])
                chairman_index.append(i)
            elif title == '董事':
                director_x.append(node_x[i])
                director_y.append(node_y[i])  
                director_index.append(i)
            elif title == '獨立董事':
                independent_director_x.append(node_x[i])
                independent_director_y.append(node_y[i])
                independent_director_index.append(i)
            elif title == '代表人':
                representative_x.append(node_x[i])
                representative_y.append(node_y[i])
                representative_index.append(i)
            elif title == '監察人':
                supervisor_x.append(node_x[i])
                supervisor_y.append(node_y[i])
                supervisor_index.append(i)
            elif title == '副董事長':
                vice_chairman_x.append(node_x[i])
                vice_chairman_y.append(node_y[i])
                vice_chairman_index.append(i)
            else:
                nothing_x.append(node_x[i])
                nothing_y.append(node_y[i])  
                nothing_index.append(i)
                
        else:
            corporate_x.append(node_x[i])
            corporate_y.append(node_y[i])
            corporate_index.append(i)

    data = [(chairman_x, chairman_y, '董事長', chairman_index), 
            (director_x, director_y, '董事', director_index),
            (independent_director_x, independent_director_y, '獨立董事', independent_director_index),
            (representative_x, representative_y, '代表人', representative_index),
            (supervisor_x, supervisor_y, '監察人', supervisor_index),
            (vice_chairman_x, vice_chairman_y, '副董事長', vice_chairman_index),
            (corporate_x, corporate_y, '公司', corporate_index),
            (nothing_x, nothing_y, '無標記', nothing_index)]

    for node_x, node_y, name, index in data:
        node_trace = go.Scatter(
            x=node_x, y=node_y,
            name=name,
            mode='markers',
            hoverinfo='text',
            marker=dict(
                colorscale='aggrnyl',
                reversescale=True,
                color='red',
                size=25,
                line_width=2))

        node_text = []
        node_color = []
        for node, adjacencies in enumerate(G.adjacency()):
            if node not in index:
                continue
            text = reverse_company_node_subset_map[adjacencies[0]]
            try:
                title = company_node_subset[company_node_subset['姓名'] == adjacencies[0]]['職稱'].values[0]
                text = '{}'.format(title)+'<br>'+text
                node_text.append(text)
            except:
                text = text
                node_text.append(text)

            try :
                node_color.append(company_node_subset[company_node_subset['姓名'] == adjacencies[0]]['顏色'].values[0])
            except:
                node_color.append('#b07b4a')

        node_trace.marker.color = node_color
        node_trace.text = node_text
        fig.add_trace(node_trace)


        #Special Case: target company
        for i,v in enumerate(index):
            if list(G.nodes())[v] == company_node_subset_map[corporate[corporate['公司名稱'] == corporate_name]['公司全稱'].values[0]]: 
                node_trace = go.Scatter(
                    x=[node_x[i]], y=[node_y[i]],
                    mode='markers',
                    name='目標檢查公司',
                    hoverinfo='text',
                    marker=dict(
                        colorscale='aggrnyl',
                        reversescale=True,
                        color='#b07b4a',
                        line=dict(
                            color='blue',
                            width=25
                            ),
                        size=30,
                        line_width=2))
                node_trace.text = corporate[corporate['公司名稱'] == corporate_name]['公司全稱'].values[0]
                fig.add_trace(node_trace)

        # Special case: target people
        for i,v in enumerate(index):
            if list(G.nodes())[v] in target_people: 
                node_trace = go.Scatter(
                    x=[node_x[i]], y=[node_y[i]],
                    mode='markers',
                    name='高風險人物',
                    hoverinfo='text',
                    marker=dict(
                        colorscale='aggrnyl',
                        reversescale=True,
                        color=company_node_subset[company_node_subset['姓名'] == list(G.nodes())[v]]['顏色'].values[0],
                        line=dict(
                            color='red',
                            width=25
                            ),
                        size=30,
                        line_width=2))
                text = reverse_company_node_subset_map[list(G.nodes())[v]]
                title = company_node_subset[company_node_subset['姓名'] == list(G.nodes())[v]]['職稱'].values[0]
                text = '{}'.format(title)+'<br>'+text
                node_trace.text = text
                fig.add_trace(node_trace)

    layout=go.Layout(
            title='董事與法人結構',
            titlefont_size=16,
            showlegend=True,
            hovermode='closest',
            margin=dict(b=20,l=5,r=5,t=40),
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False))
    fig.update_layout(layout)
    fig.write_html("./visual_output/governance/director_structure/{}_director_structure.html".format(corporate_name))
    #fig.show()


    # ## Market

    # In[31]:


    c = list(corporate['公司名稱'])
    c[16] = '遊戲橘子'
    df_chinatime = pd.DataFrame()
    df_tvbs = pd.DataFrame()
    df_setn = pd.DataFrame()

    for l in c:
        try:
            d_temp = pd.read_csv('./data/crawler/中時新聞/{}中時.csv'.format(l), header=None)
            d_temp['company'] = l
            df_chinatime = df_chinatime.append(d_temp, ignore_index=True)
        except:
            continue
            
    for l in c:
        try:
            d_temp = pd.read_csv('./data/crawler/tvbs新聞/{}tvbs.csv'.format(l), header=None)
            d_temp['company'] = l
            df_tvbs = df_tvbs.append(d_temp, ignore_index=True)
        except:
            continue
            
    for l in c:
        try:
            d_temp = pd.read_excel('./data/crawler/三立新聞爬蟲/{}.SETN.xlsx'.format(l))
            d_temp['company'] = l
            df_setn = df_setn.append(d_temp, ignore_index=True)
        except:
            continue


    # In[32]:


    # Clean Data
    df_tvbs['date'] = df_tvbs[0].apply(lambda x : ''.join(x[:int(len(x)/2)].split('/')[:2]))
    df_chinatime['time'] = df_chinatime[0].apply(lambda x : x.split('\n')[1][:5])
    df_chinatime['date'] = df_chinatime[0].apply(lambda x : ''.join(x.split('\n')[1][5:].split('/')[:2]))
    df_setn['date'] = df_setn['TIME'].apply(lambda x: ''.join(x.split('/')[:2]))

    df_tvbs[2] = df_tvbs[2].apply(lambda x: x.replace('最HOT話題在這！想跟上時事，快點我加入TVBS新聞LINE好友！\n◤疫情再升溫 防疫不出門◢👉還有缺什麼嗎？到防疫專區GO！👉宅在家不怕腫！睡好代謝好靠這個👉百大百貨品牌momo都有👉關鍵時刻！熬雞精提升保護力👉限時搶購\u3000超多品牌爆殺倒數👉高中進度攻略班，防疫限定免費 ', ''))
    df_tvbs[2] = df_tvbs[2].apply(lambda x: x.replace('中央社）\xa0～開啟小鈴鐺\u3000TVBS YouTube頻道新聞搶先看\u3000快點我按讚訂閱～', ''))


    # In[33]:


    # Concat data
    PUNCTS = [',', '.', '"', ':', ')', '(', '-', '!', '?', '|', ';', "'", '$', '&', '/', '[', ']', '>', '%', '=', '#', '*', '+', '\\', '•',  '~', '@', '£', 
              '·', '_', '{', '}', '©', '^', '®', '`',  '<', '→', '°', '€', '™', '›',  '♥', '←', '×', '§', '″', '′', 'Â', '█', '½', 'à', '…', 
              '“', '★', '”', '–', '●', 'â', '►', '−', '¢', '²', '¬', '░', '¶', '↑', '±', '¿', '▾', '═', '¦', '║', '―', '¥', '▓', '—', '‹', '─', 
              '▒', '：', '¼', '⊕', '▼', '▪', '†', '■', '’', '▀', '¨', '▄', '♫', '☆', 'é', '¯', '♦', '¤', '▲', 'è', '¸', '¾', 'Ã', '⋅', '‘', '∞', 
              '∙', '）', '↓', '、', '│', '！', '（', '»', ' ', 'LINE', 'HOT','TVBS','？','\u3000', '「', '」', '，', '♪', '\xa0', '╩', '，', '╚', '³', '・', '╦', '╣', '╔', '╗', '▬', '❤', 'ï', 'Ø', '¹', '≤', '‡', '√', '#', '。','—–', '👉', '        ', '\n']

    def clean_puncts(sentences):
        sentences = [[w for w in s if (not w in PUNCTS) & (len(w)>1)] for s in sentences]
        return sentences

    with open('./data/sw.txt') as f :
        sw = f.readlines()
    sw = [i.split('\n')[0] for i in sw]

    jieba.set_dictionary('./data/dict.txt')

    cut = []
    for i in df_chinatime[df_chinatime['company'] == corporate_name][2]:
        cut.append(jieba.analyse.extract_tags(i, topK=30, withWeight=False, allowPOS=()))
        
    chinatime_cut = clean_puncts(cut)
    chinatime_cut = [[w for w in s if not w in sw] for s in chinatime_cut]
    l1 = sum(chinatime_cut, [])

    cut = []
    for i in df_tvbs[df_tvbs['company'] == corporate_name][2]:
        cut.append(jieba.analyse.extract_tags(i, topK=30, withWeight=False, allowPOS=()))
        
    tvbs_cut = clean_puncts(cut)
    tvbs_cut = [[w for w in s if not w in sw] for s in tvbs_cut]
    l2 = sum(tvbs_cut, [])

    cut = []
    for i in df_setn[df_tvbs['company'] == corporate_name]['Content']:
        cut.append(jieba.analyse.extract_tags(i, topK=30, withWeight=False, allowPOS=()))
        
    setn_cut = clean_puncts(cut)
    setn_cut = [[w for w in s if not w in sw] for s in setn_cut]
    l3 = sum(setn_cut, [])

    l = []
    l.extend(l1)
    l.extend(l2)
    l.extend(l3)

    if len(l) == 0:
        l = ['無新聞數據']
        
    df = pd.DataFrame(Counter(l).most_common())


    # In[34]:


    plt.figure(figsize=(24,16))
    wc = WordCloud(width=1600, height=800, background_color='white', font_path='/Users/howardchung/PPT/字體/TaipeiSansTCBeta-Light.ttf') 
    wc.generate(' '.join(l))
    plt.imshow(wc)
    plt.axis("off")
    plt.tight_layout(pad=0)
    plt.savefig("./visual_output/market/word_cloud/{}_word_cloud.png".format(corporate_name))
    #plt.show()


    # In[35]:


    # Amount of news

    new_df = pd.DataFrame(df_tvbs[df_tvbs['company'] == corporate_name]['date'].append(df_setn[df_setn['company'] == corporate_name]['date']).append(df_chinatime[df_chinatime['company'] == corporate_name]['date'])).reset_index(drop=True)
    new_df = pd.DataFrame(new_df.value_counts()).reset_index()
        
    draw_df = pd.DataFrame()
    for y in list(set([d[:4] for d in list(new_df['date'])])):
        d_temp = new_df[new_df['date'].str.contains(y)]
        for q, l in [(1,['01', '02', '03']), (2, ['04', '05', '06']), (3,['07', '08', '09']), (4, ['10', '11', '12'])]:
            c = 0
            for m in l:
                try:
                    c += d_temp[d_temp['date'] == '{}{}'.format(y, m)][0].values[0]
                except:
                    continue
            if c !=0:
                draw_df = draw_df.append({'year':'{}Q{}'.format(y, q), 'count':c}, ignore_index=True)

    if len(new_df) == 0:
        draw_df = pd.DataFrame({'year':[0], 'count':[0]})
        
    fig = go.Figure()
    fig.add_trace(go.Bar(x=draw_df['year'], y=draw_df['count'], marker_color='red', showlegend=True))

    fig.update_layout(
        barmode='stack', 
        title='提及新聞-新聞數量統計',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='新聞數量',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            categoryorder='category ascending',
            title='年份',
            titlefont_size=16,
            tickfont_size=14,
        ),
        showlegend=False,
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.update_traces(marker_color='rgb(55, 83, 109)')
    fig.write_html("./visual_output/market/news_statistic/{}_news_statistic.html".format(corporate_name))
    #fig.show()


    # In[36]:


    draw_df = pd.DataFrame({'新聞台':['TVBS', '三立新聞台', '中國時報'],
                          '數量': [len(df_tvbs[df_tvbs['company'] == corporate_name]), 
                                 len(df_setn[df_setn['company'] == corporate_name]), 
                                 len(df_chinatime[df_chinatime['company'] == corporate_name])]})

    fig = go.Figure()
    fig.add_trace(go.Bar(x=draw_df['新聞台'], y=draw_df['數量'], marker_color='red', showlegend=True))

    fig.update_layout(
        barmode='stack', 
        title='提及新聞-新聞台數量統計',
        xaxis_tickfont_size=14,
        yaxis=dict(
            title='新聞數量',
            titlefont_size=16,
            tickfont_size=14,
        ),
        xaxis=dict(
            categoryorder='category ascending',
            title='新聞台',
            titlefont_size=16,
            tickfont_size=14,
        ),
        showlegend=False,
        bargap=0.15, # gap between bars of adjacent location coordinates.
        bargroupgap=0.1 # gap between bars of the same location coordinate.
    )
    fig.update_traces(marker_color='rgb(55, 83, 109)')
    fig.write_html("./visual_output/market/news_desk_statistic/{}_news_desk_statistic.html".format(corporate_name))
    #fig.show()


    # In[41]:


    score = score.append(score_dict, ignore_index=True)
score.to_csv('score.csv', index=False)


# In[ ]:




