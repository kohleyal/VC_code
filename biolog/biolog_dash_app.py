
import pandas as pd
import plotly.express as px
#from jupyter_dash import JupyterDash for working in a jupter notebook. need to change app = Dash(__name__) to app = JupyterDash(__name__)
from dash import dcc, html, Dash
from dash.dependencies import Input, Output, State
from tkinter import filedialog
from webbrowser import open as openbroswer

app = Dash(__name__)
'''class for each excel file that will extract all the data on each sheet and make an instance with all the sheets run through pandas data transformation
    and ready to make graphs with.'''

biolog_plate_dic = {'1' : 'Carbon Sources', '2A' : 'Carbon Sources', '3B' : 'Nirtogen Sources', '4A' : 'Poshporus and Sulfur Sources',
                    '5' : 'Nutrient Supplements', '6' : 'Peptide Nitrogen Sources', '7' : 'Peptide Nitrogen Sources', '8' : 'Peptide Nitrogen Sources',
                    '9' : 'Osmolytes', '10' : 'pH', '21D' : 'Chemical Sensitivity', '22D' : 'Chemical Sensitivity', '23A' : 'Chemical Sensitivity',
                    '24C' : 'Chemical Sensitivity', '25D' : 'Chemical Sensitivity' }

class Biolog_excel_file():
  ''' hardcoded values to work with specific internal excel file.'''
    header_row_absorbance = 71 #pandas starts at 0 and excel starts at 1. row on excel sheet will be 72
    header_row_hours = 71
    data_column_range = 'A:N'
    hours_column = 'S'
    index_of_timelabel_0 = 7
    number_of_timepoints = 7
    sheet_names = [f'{x}' for x in range(1,23)]

    def __init__(self, excel_file, title = None):
        self.excel_file = excel_file
        if not title:
            title = excel_file[excel_file.rfind('/') + 1 : excel_file.rfind('.')]
        else:
            title = excel_file[excel_file.rfind('/') + 1 : excel_file.rfind('.')] + ' - Copy'

        self.file_name = title
        self.title = f'{title}'
        self.plate_info = []
        self.dic_of_df_plates = pd.read_excel(self.excel_file, self.sheet_names, header = self.header_row_absorbance, usecols = self.data_column_range)
        self.dic_of_df_hours = pd.read_excel(self.excel_file, self.sheet_names, header= self.header_row_hours, usecols = self.hours_column, nrows= self.number_of_timepoints)
        for sheetname in self.sheet_names:
            '''build list of dictionaries for the drop down menu. with the plate and sugar info to be displayed'''
            plate_info = self.dic_of_df_plates.get(sheetname)
            try:
                plate_info = plate_info[['Plate ', 'Sugar']].loc[0, :].values.tolist()
            except KeyError:
                plate_info = plate_info[['Plate', 'Sugar']].loc[0, :].values.tolist()

            plate_info = [f'{x}' for x in plate_info]
            if plate_info[0] in biolog_plate_dic:
                plate_info[0] = biolog_plate_dic.get(plate_info[0])
            self.plate_info.append({'label' : f'{sheetname}. {" - ".join(plate_info)}', 'value' : sheetname })

            '''build df to use with plotly and assign it to an attribute based on the sheet name'''
            setattr(self, f'sheet{sheetname}', self.excel_sheet_data_proccessing(sheetname))
        

    def excel_sheet_data_proccessing(self, sheet_key):
        df_plate = self.dic_of_df_plates.get(sheet_key)
        df_hours = self.dic_of_df_hours.get(sheet_key)

        #this pulls apart the letter number location data and organizes it so they will be entered in the right order in the figure
        df_plate['loca_letters'] = df_plate['Location'].str[:1]
        df_plate['loca_numbers'] = df_plate['Location'].str[1:].astype(float)
        df_plate['file'] = self.title

        '''Have to make and use component location beause if component is used some plates have duplicates of componetes that
        will be aggregated. I dont want to use location because the titles are then just well locations. could add in a dictionary
        and override with plotly's labels inside of px.line '''
        df_plate['component_location'] = df_plate['Component'] + ' ' + df_plate['Location']
        

        #get the sampling timepoints. 
        timepoint_header_list = df_plate.columns[self.index_of_timelabel_0 : self.index_of_timelabel_0 + self.number_of_timepoints].tolist()
        df_hours['timepoint_name'] = timepoint_header_list
        
        '''for melt id_vars used component, location, and componet_location because some componets are repeated and
            the melt will compile the repeating componets if only component is used. 
            Then replaces the timepoint lables with actual time integers. '''
        
        df_plate_melt = (df_plate.melt(id_vars=['Component', 'Location', 'component_location', 'file','loca_letters', 'loca_numbers'], value_vars= timepoint_header_list, var_name='hour', value_name='abs')
                                .replace(df_hours.iloc[:,1].tolist(),df_hours.iloc[:,0]))
        return df_plate_melt




experiment_one_file_str = filedialog.askopenfilename(filetypes= [('Excel file','.xlsx'), ('Any','.*')]) 

experiment_two_file_str = filedialog.askopenfilename(filetypes= [('Excel file','.xlsx'), ('Any','.*')]) 



experiment_one = Biolog_excel_file(experiment_one_file_str)

'''for right now. a quick way to look at the same file. if the titles are the same, they will be graghed over each other.'''
if experiment_one_file_str == experiment_two_file_str:
    experiment_two = Biolog_excel_file(experiment_two_file_str, f'{experiment_two_file_str} - copy')
else: 
    experiment_two = Biolog_excel_file(experiment_two_file_str)




'''build the layout for the dash web page'''
app.layout = html.Div([

    html.H2("Biolog Data Interface", style={'text-align' : 'center'}),

    html.Div(children=[
        html.Div([
            html.Label(f'{experiment_one.title}', 'drop1_label'),
            dcc.Dropdown(experiment_one.plate_info, id='biolog_excel_sheet1', value='1', clearable=False)],
                style={'padding': 5, 'flex':1, 'display': 'flex', 'flex-direction': 'column'}),
        
        html.Div([
            html.Label(f'{experiment_two.title}', 'drop2_label'),
            dcc.Dropdown(experiment_two.plate_info, id='biolog_excel_sheet2', value='2', clearable=False)],
                style={'padding': 5, 'flex':1, 'display': 'flex', 'flex-direction': 'column'}),

        html.Button(id='submit_button', n_clicks=0, children='Submit', style={ 'flex-shrink': 2, 'flex': 1, 'margin':r'20px 50px 2px 2px', 'padding': '9px 0px 9px 0px'})],
        
        style={'display': 'flex', 'flex-direction': 'row', 'align-items': 'center'}),

    html.H4(id='label', children='''Select one or two files to display.'''),
    dcc.Checklist([{'label': f'{experiment_one.file_name}', 'value' : 1}, {'label' : f'{experiment_two.file_name}', 'value' : 2}],
                    [1, 2], id='file_checklist', inline = True, inputStyle={"margin-left": "20px"}),
    html.Br(),

    dcc.Graph(id='plate_figure', figure={}) 
], style={'display': 'flex', 'flex-direction': 'column'})




'''function to make the data and the figure to put into the web page. should pull apart into figure maker and data arranger'''

'''need to differentiate between experiment one and two. either have it send over the object or have two different app callback functions. one with each experiment object'''
@app.callback(
    Output(component_id='plate_figure', component_property='figure'),
    Input(component_id='submit_button', component_property='n_clicks', ),
    State(component_id='biolog_excel_sheet1', component_property= 'value'),
    State(component_id='biolog_excel_sheet2', component_property= 'value'),
    State('file_checklist', 'value'),
    prevent_initial_call=True
)
def make_figure(bttn, plate1_df_attr_name, plate2_df_attr_name, files_used_checklist):
    if len(files_used_checklist) > 1:
        newfile = (pd.concat([getattr(experiment_two, f'sheet{plate2_df_attr_name}'),getattr(experiment_one, f'sheet{plate1_df_attr_name}')])
                    .sort_values(by=['loca_letters', 'loca_numbers', 'hour']))
        
    elif files_used_checklist[0] == 2:
        newfile = getattr(experiment_two, f'sheet{plate2_df_attr_name}').sort_values(by=['loca_letters', 'loca_numbers', 'hour'])

    else: 
        newfile = getattr(experiment_one, f'sheet{plate1_df_attr_name}').sort_values(by=['loca_letters', 'loca_numbers', 'hour'])

    fig = px.line(newfile,
        x='hour', y='abs', hover_name='Location',
        facet_col='component_location', facet_col_wrap=12,
        facet_col_spacing=0.008, facet_row_spacing=0.025,
        height=900, width=1600, markers=True, color='file')
        # symbol='series')

    #font size for the charts and and removing the = in the titles.
    for anno in fig['layout']['annotations']:
        anno['font']['size']= 8.5
        anno['text'] = anno['text'].split('=')[-1]

    #make the window for all the outputs the same range. Many want to modify off of range for all sheets or not keep.
    (fig.update_traces(marker={'size': 3})
        .update_xaxes(title_font_size=12)
        .update_yaxes(title_font_size=12)
        .update_layout(legend=dict(
            orientation="h", yanchor="bottom", y=1.03, xanchor="right", x=0.5))
    )
    return fig


print('once')
if __name__ == '__main__':
    openbroswer('http://127.0.0.1:8050/', 1) #opens browser window with defualt dash app path
    app.run_server(debug=False)

    



