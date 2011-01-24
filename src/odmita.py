# coding=utf-8
'''
Created on 24.01.2011

@author: Mr.Hant@gmail.com
'''

import yapgvb, docx, os
from docx import table, heading, paragraph, picture, search, replace, pagebreak
if __name__ == '__main__':        
    
    A = [[0,1,1,1,0,1,],
         [1,0,1,0,0,1,],
         [1,1,0,1,1,1,],
         [1,0,1,0,0,1,],
         [0,0,1,0,0,1,],
         [1,1,1,1,1,0,]]
    
    graph = yapgvb.Graph("my_graph")
    
    node=[]
    # Number of verticies
    n = len(A)
    # Number of edges 
    r = 0

    # Create nodes
    for i in range(n):
        node.append(graph.add_node())
        node[i].label="X"+str(i+1)
 
    # Draw a basic graph image 
    for i in range(n):
        for j in range(i,n):
            if (A[i][j]==1):
                A[j][i]=1 # if edge exists => add backward mark in A matrix
                r+=1 # increase number of edges
                graph.add_edge(node[i],node[j]) # add an adge for drawing
    graph.layout(yapgvb.engines.dot) # 'Dot' fits the best
    format = yapgvb.formats.png
    filename = 'graph.%s' % format
    print "  Rendering %s ..." % filename
    graph.render(filename) # Draws a graph

    relationships = docx.relationshiplist()
    document = docx.newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=docx.nsprefixes)[0]
   
    docbody.append(heading(u'''6. Найти инварианты графа, заданного матрицей смежности ''',1)  )   
    
    # Create and insert basic table containing input info
    Table_Basic=[]
    for i in range(n+1):
        Table_Basic.append([])
        for j in range(n+1):
            Table_Basic[i].append("")
    for i in range(1,n+1):
        Table_Basic[0][i] = "x"+str(i)
        Table_Basic[i][0] = "x"+str(i)
    for i in range(1,n+1):
        for j in range(1, n+1):
            Table_Basic[i][j]=str(A[i-1][j-1])        
    docbody.append(table(Table_Basic,False,borders={'all':{'color':'auto','space':1,'sz':1}}))
       
    docbody.append(heading(u'Решение:',2))
    docbody.append(paragraph(u'Согласно отношениям смежности, изобразим граф:'))
    
    # Insert an image of graph
    relationships,picpara = picture(relationships,'graph.png','',200,200)
    docbody.append(picpara)

    docbody.append(heading(u'1. Количество вершин  n = %i'%(n),2))
    
    docbody.append(heading(u'2. Количество ребер r = %i'%(r),2))
    
    f = 2-n+r
    docbody.append(heading(u'3. Количество граней f = %i'%(f),2))
    docbody.append(paragraph('n-r+f = 2'))
    docbody.append(paragraph('f = 2-n+r = 2-6+11 = 7'))
    
    
    """
    ## Some examples of working with docx module
    ## TODO: Remove examples after finishing work
    
    # Add a numbered list
    for point in ['''COM automation''','''.net or Java''','''Automating OpenOffice or MS Office''']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(paragraph('''For those of us who prefer something simpler, I made docx.''')) 
    
    docbody.append(heading('Making documents',2))
    docbody.append(paragraph('''The docx module has the following features:'''))

    # Add some bullets
    for point in ['Paragraphs','Bullets','Numbered lists','Multiple levels of headings','Tables','Document Properties']:
        docbody.append(paragraph(point,style='ListBullet'))

    docbody.append(paragraph('Tables are just lists of lists, like this:'))
    # Append a table
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))

    docbody.append(heading('Editing documents',2))
    docbody.append(paragraph('Thanks to the awesomeness of the lxml module, we can:'))
    for point in ['Search and replace','Extract plain text of document','Add and delete items anywhere within the document']:
        docbody.append(paragraph(point,style='ListBullet'))
 
    # Search and replace
    print 'Searching for something in a paragraph ...',
    if search(docbody, 'the awesomeness'): print 'found it!'
    else: print 'nope.'
    
    print 'Searching for something in a heading ...',
    if search(docbody, '200 lines'): print 'found it!'
    else: print 'nope.'
    
    print 'Replacing ...',
    docbody = replace(docbody,'the awesomeness','the goshdarned awesomeness') 
    print 'done.'

    # Add a pagebreak
    docbody.append(pagebreak(type='page', orient='portrait'))

    docbody.append(heading('Ideas? Questions? Want to contribute?',2))
    docbody.append(paragraph('''Email <python.docx@librelist.com>'''))
    """

    # Create our properties, contenttypes, and other support files
    coreprops = docx.coreproperties(title=u'Контрольная работа',subject=u'ОДМиТА, Задание №6',creator='Mr.Hant@gmail.com',keywords=['БГУИР',u'ОДМиТА','python'])
    appprops = docx.appproperties()
    contenttypes = docx.contenttypes()
    websettings = docx.websettings()
    wordrelationships = docx.wordrelationships(relationships)
    
    # Save our document
    docx.savedocx(document,coreprops,appprops,contenttypes,websettings,wordrelationships,'./output/solution.docx')
    # Remove temporary image files
    os.remove('graph.png')
    
    
    