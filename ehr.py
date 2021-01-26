import openpyxl as pxl
import re
import igraph
from igraph import *
import sys
from treelib import Node, Tree
import secrets
  
# the setrecursionlimit function is 
# used to modify the default recursion 
# limit set by python. Using this,  
# we can increase the recursion limit 
# to satisfy our needs 
  
#sys.setrecursionlimit(10**6) 

# separated by +-*/
# function: name(a,b,c) arguments comma separated
# SUM| function: name(a:b) series defined by a:b
def make_annotations(M,pos, text, font_size=10, font_color='rgb(250,250,250)'):
    L=len(pos)
    if len(text)!=L:
        raise ValueError('The lists pos and text must have the same len')
    annotations = []
    for k in range(L):
        annotations.append(
            dict(
                text=text[k], # or replace labels with a different list for the text within the circle
                x=pos[k][0], y=2*M-pos[k][1],
                xref='x1', yref='y1',
                font=dict(color=font_color, size=font_size),
                showarrow=False)
        )
    return annotations

def create_graph(G):
    nr_vertices = len(G.vs)
    v_label = G.vs["label"]#list(map(str, range(nr_vertices)))
    lay = G.layout_reingold_tilford(mode="out", root=[1], rootlevel=[3])
    position = {k: lay[k] for k in range(nr_vertices)}
    Y = [lay[k][1] for k in range(nr_vertices)]
    M = max(Y)

    es = EdgeSeq(G) # sequence of edges
    E = [e.tuple for e in G.es] # list of edges

    L = len(position)
    Xn = [position[k][0] for k in range(L)]
    Yn = [2*M-position[k][1] for k in range(L)]
    Xe = []
    Ye = []
    for edge in E:
        Xe+=[position[edge[0]][0],position[edge[1]][0], None]
        Ye+=[2*M-position[edge[0]][1],2*M-position[edge[1]][1], None]

    labels = G.vs["label"]
    names = G.vs["name"]


    import plotly.graph_objects as go
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=Xe,
                    y=Ye,
                    mode='lines',
                    line=dict(color='rgb(210,210,210)', width=1),
                    hoverinfo='none'
                    ))
    fig.add_trace(go.Scatter(x=Xn,
                    y=Yn,
                    mode='markers',
                    name='bla',
                    marker=dict(symbol='circle-dot',
                                    size=40,
                                    color='#6175c1',    #'#DB4551',
                                    line=dict(color='rgb(50,50,50)',
                                     width=1,
                                     )
                                    ),
                    text=names,
                    hoverinfo='text',
                    opacity=0.8
                    ))
    axis = dict(showline=False, # hide axis line, grid, ticklabels and  title
            zeroline=False,
            showgrid=False,
            showticklabels=False,
            )

    fig.update_layout(title= 'EXCEL COMPUTATION HIERARCHY',
                annotations=make_annotations(M,position, v_label),
                font_size=12,
                showlegend=False,
                xaxis=axis,
                yaxis=axis,
                margin=dict(l=40, r=40, b=85, t=100),
                hovermode='closest',
                plot_bgcolor='rgb(248,248,248)',
                width=15000,
                height=4000
                )
    fig.show()

def get_cells(query):
    #print("DEBUG:")
    #print(query)
    flist = []
    if type(query) == tuple:
        for cs in query:
            if(type(cs) == tuple):
                flist.append(cs[0])
            else:
                flist.append(cs)
    else:
        flist.append(query)

    cell_list = []
    for cs in flist:
        val = cs.value
        val = str(val)
        if(val != None):
            val = re.sub(r"[$=]", "", val)
            res = re.findall(r'((?:\w+!)?C\d+(?::C\d+)?)+', val)
            if(len(res)):
                cell_list.append(res)
    return cell_list
    
#globally referenced graph object        
G = Graph()
T = Tree()
def generate_hierarchy_r(root, sheet, vr, cell_dict):

    #graph
    if(type(vr) != str):
        G.add_vertices(1)
        vc = G.vs[-1]
        vc["label"]=root
        try:
            vc["name"]=sheet + '!' + root + wb[sheet][root].value
        except:
            vc["name"]=sheet + '!' + root
        G.add_edges([(vr, vc)])
    #Tree
    else: 
        vc = secrets.token_hex(5)
        T.create_node(sheet + '!' + root, vc, parent=vr)
    #print(tabs+sheet + "!" + root)
    if sheet+'!'+root in cell_dict:
        return {}
    cell_dict[sheet+'!'+root] = wb[sheet][root]
    cell_lists = get_cells(wb[sheet][root])
    
    for cell_list in cell_lists:
        for cell in cell_list:
            if(re.match(r'\w+!\w+', cell)):
                snr = re.split('!', cell)
                cell = snr[1]
                cell_dict = {**cell_dict, **generate_hierarchy_r(cell, snr[0], vc, cell_dict)}
            else:
                cell_dict = {**cell_dict, **generate_hierarchy_r(cell, sheet, vc, cell_dict)}
    return cell_dict
wb = pxl.load_workbook('CLNsPricer.xlsm')
wb.active = wb['CLNs']

def generate_hierarchy(root, sheet, display):
    if display == 'graph':
        G.add_vertices(1)
        vr = G.vs[-1]
        cell_dict = generate_hierarchy_r(root, sheet, vr, {})
        create_graph(G)
    else:
        vr = secrets.token_hex(5)
        T.create_node(sheet + '!' + root, vr)
        cell_dict = generate_hierarchy_r(root, sheet, vr, {})
        T.show()


generate_hierarchy('C100', 'CLNs', 'graph')
generate_hierarchy('C100', 'CLNs', 'tree')




