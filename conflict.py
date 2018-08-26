import  xlrd
import networkx as nx
import matplotlib.pyplot as plt
from xlrd import open_workbook

##############################33
def cycle_exists(G):  # - G is a directed graph
    color = {u: "white" for u in G}  # - All nodes are initially white
    found_cycle = [False]  # - Define found_cycle as a list so we can change
    
    for u in G:  # - Visit all nodes.
        if color[u] == "white":
            dfs_visit(G, u, color, found_cycle)
        if found_cycle[0]:
            break

    if found_cycle[0]==False :
        print("given schedule is *******************  CONFLICT***************")
        #print(nx.topological_sort(G,nbunch=None,reverse=False))
        print(nx.topological_sort_recursive(G,nbunch=None,reverse=False))

    else:
        print("cycle found so ... -->>")
        print("given schedule is ******************* NOT CONFLICT ***************")

    return found_cycle[0]


# -------

def dfs_visit(G, u, color, found_cycle):
    if found_cycle[0]:  # - Stop dfs if cycle is found.
        return
    color[u] = "gray"  # - Gray nodes are in the current path
    for v in G[u]:  # - Check neighbors, where G[u] is the adjacency list of u.
        if color[v] == "gray":  # - Case where a loop in the current path is present.
            found_cycle[0] = True
            return
        if color[v] == "white":  # - Call dfs_visit recursively.
            dfs_visit(G, v, color, found_cycle)
    color[u] = "black"  # - Mark node as done.

######################################
################## final nodes relation
#####################################

def store_node(lst,c,lastRow_no,noOfRows,noOfCols,lists):
    node_no = c
    c-=1
    if not lst:
        pass
    else:
        lists[c]=list(set(lists[c])|set(lst))

    if lastRow_no==noOfRows :
        print("final node list which will be connected to each other index+1 to list[]")
        print('***************************************************')
        print(lists)
        print('***************************************************')
        G = nx.DiGraph()

        #################
        #### draw no. of nodes

        for j in range(noOfCols):
            G.add_node(j + 1)


        ################
        #### seprate lists as nodes

        j=0
        for i in range(noOfCols):
            g=chr(75+i)
            g=lists[i]
            j += 1
            for x in range(len(g)):
                y = g[x]
                G.add_edge(j, y)
        nx.draw(G, with_labels=True)
        #plt.savefig("simple_path.png")  # save as png
        plt.show()  # display
        cycle_exists(G)
######################################
################## arrow from A to list or variable
#####################################

def A_On_Arrow(firstRow_value,r,c,lastRow_no,noOfRows,noOfCols,lists):
    A_list=[]
    if firstRow_value=='r(A)':
        for sheet in book.sheets():
            for rowidx in range(sheet.nrows):
                row = sheet.row(rowidx)
                for colidx, cell in enumerate(row):
                    if cell.value == 'w(A)' and rowidx > r :
                        if(colidx != c ):
                            A_list.append(colidx)

    elif firstRow_value == 'w(A)':
        for sheet in book.sheets():
            for rowidx in range(sheet.nrows):
                row = sheet.row(rowidx)
                for colidx, cell in enumerate(row):
                    if cell.value == 'r(A)' and rowidx > r :
                        if colidx != c:
                            A_list.append(colidx)

                    elif cell.value == 'w(A)' and rowidx > r :
                        if (colidx != c):
                            A_list.append(colidx)

    store_node(A_list,c,lastRow_no,noOfRows,noOfCols,lists)


######################################
################## arrow from B to list or variable
#####################################

def B_On_Arrow(item,r,c,lastRow_no,noOfRows,noOfCols,lists):
    B_list=[]
    if firstRow_value=='r(B)':
        for sheet in book.sheets():
            for rowidx in range(sheet.nrows):
                row = sheet.row(rowidx)
                for colidx, cell in enumerate(row):
                    if cell.value == 'r(B)' and rowidx > r:
                        if colidx != r :
                            print(rowidx, colidx)

                    elif cell.value == 'w(B)' and rowidx > r:
                        if(colidx != r ):
                            B_list.append(colidx)
    elif firstRow_value=='w(B)':
        for sheet in book.sheets():
            for rowidx in range(sheet.nrows):
                row = sheet.row(rowidx)
                for colidx, cell in enumerate(row):
                    if cell.value == 'r(B)' and rowidx > r:
                        if(colidx!=r):
                            B_list.append(colidx)

                    elif cell.value == 'w(B)' and rowidx > r:
                        if colidx != r :
                            B_list.append(colidx)
    store_node(B_list,c,lastRow_no,noOfRows,noOfCols,lists)

######################################
################## function call as variable name
#####################################

def callingfunction_list(firstRow_value,r,c,lastRow_no,noOfRows,noOfCols,lists):

    if firstRow_value[2]=='A':
        A_On_Arrow(firstRow_value,r,c,lastRow_no,noOfRows,noOfCols,lists)
    elif firstRow_value[2]=='B':
        B_On_Arrow(firstRow_value, r, c,lastRow_no,noOfRows,noOfCols,lists)

######################################
################## first row value calculated to find which variable
#####################################

book = open_workbook('demo.xlsx')
for sheet in book.sheets():
    lastRow_no=0
    for noOfRows in range(sheet.nrows):
        row = sheet.row(noOfRows)
        for noOfCols, cell in enumerate(row):
            pass
    lists = [[]] * noOfCols
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            lastRow_no=rowidx
            if rowidx>0 and cell.value !='':
                firstRow_value = str(cell.value)
                callingfunction_list(firstRow_value,rowidx,colidx,lastRow_no,noOfRows,noOfCols,lists)

