import plotly.graph_objects as go
import pandas as pd
import locale

locale.setlocale(locale.LC_ALL, '')

filepath = 'C:/Users/marce/Documents/Python/'

# READ IN DATA
adv = pd.read_excel(filepath + 'SWD/Which Would You Choose/Quarterly Advertising.xlsx',
                    sheet_name='Original Data')
adv['diff'] = adv['Q3Spend'] - adv['Q2Spend']

fig = go.Figure()

for i in range(0, len(adv)):
    if adv['diff'].iloc[i] > 0:
        colordiff = 'green'
    else:
        colordiff = 'rgb(171, 21, 10)'

    fig.add_shape(type='line',
                  x0=adv['Q2Spend'].iloc[i],
                  y0=adv['category'].iloc[i],
                  x1=adv['Q3Spend'].iloc[i],
                  y1=adv['category'].iloc[i],
                  line=dict(color=colordiff, width=2),
                  xref='x',
                  yref='y',
                  layer='below'
                  )

fig.add_trace(go.Scatter(x=adv['Q2Spend'], y=adv['category'], mode='markers', name='Q2',
                         marker=dict(color='rgb(135,206,235)', size=8, line=dict(color='grey', width=2))))
fig.add_trace(go.Scatter(x=adv['Q3Spend'], y=adv['category'], mode='markers', name='Q3',
                         marker=dict(color='rgb(239,228,176)', size=8, line=dict(color='grey', width=2))))

fig.update_layout(showlegend=True, plot_bgcolor='rgb(242,242,242)', paper_bgcolor='rgb(242,242,242)',
                  yaxis=dict(autorange="reversed"),
                  xaxis_tickprefix='$', xaxis_ticksuffix='M',
                  legend=dict(orientation="h", y=1.1, x=0.43, font=dict(size=14)),
                  title_x=0.5,
                  title=dict(text='<b>Quarterly Advertising Spend</b>',
                             font=dict(size=24)))
fig.add_annotation(x=7, y=-0.15,
                   text="increase over time",
                   showarrow=False,
                   yshift=20)

fig.add_annotation(x=10, y=-0.15,
                   text="decrease over time",
                   showarrow=False,
                   yshift=20)

fig.add_shape(type='line', x0=5.5, y0=-0.62, x1=6, y1=-0.62, line=dict(color='green', width=2), xref='x', yref='y')
fig.add_shape(type='line', x0=8.5, y0=-0.62, x1=9, y1=-0.62, line=dict(color='rgb(171, 21, 10)', width=2), xref='x', yref='y')

fig.update_xaxes(showgrid=False, zeroline=False)
fig.update_yaxes(showgrid=False)

fig.update_traces(hovertemplate="%{y}" + ': ' "%{x}<extra></extra>")

fig.write_html(filepath + 'Dot Plot Example.html')
