from docx import Document
from docx.shared import Inches

import numpy as np
import pandas as pd
import collections
import matplotlib.pyplot as plt
import textdistance
import statistics 
import igraph

def processSentence(sentence:str)->str:
    """
    Deletion of some characters in the sentence

    Del emoji, and special characters, change flag emoji to country name.

    :param str sentence: input sentence
    :returns: new sentence without some special characters
    :rtype: string
    """
    caraToDel="🐝🚲⛺🌻🌲⭐⚽👨🏼‍🏫🕵🧭🇯🇻🏼‍😃👨🏻‍🏫🌎📲🏛🗽🌊🚢"  #del som emoji
    sentence=sentence.lower()
    transTable = sentence.maketrans("éèà-'.ôç≠ę:î", "eea   oc e i", "!#$%^&*()\"’‘«»,，@_/|"+caraToDel)
    sentence = sentence.translate(transTable)
    #translate some emoji
    for drap, txt in zip(["🇫🇷", "🇨🇳", "🇹🇼","🇰🇷","🇰🇵","🇩🇪","🇺🇲","🇪🇺", "🇷🇺", "🇮🇷","🇺🇳"], ["france","china", "taiwan", "south korea", "north korea", "germany", "usa", "europe", "russia", "iran","un"]):
        if(drap in sentence):
            sentence=sentence[:sentence.index(drap)]+" "+txt+" "+sentence[sentence.index(drap)+2:]
    while "\n" in sentence:
        sentence=sentence[:sentence.index("\n")]+" "+sentence[sentence.index("\n")+1:]
    while "  " in sentence:
        sentence=sentence[:sentence.index("  ")]+" "+sentence[sentence.index("  ")+2:]
    return sentence

def processListSentences(L:list):
    """
    Change a list of input sentences to a list clean sentences with :func:`processSentence` fuction

    :param list<str> L: Sentences list

    :returns: clean sentences
    :rtype: list
    """
    L_phrase=[]
    for l in L:
        a=processSentence(str(l))
        L_phrase.append(a)
    return L_phrase

def most_commun_word(L:list, min_word_size:int=3, process_word:bool=False):
    """
    :param list<str> L: Sentences list
    :param int min_word_size: minimum size of the words to be taken into consideration
    :param bool process_word: if the sentence has already been processed in :func:`processSentence`

    :returns: Return most commun word form a sentences list
    :rtype: dict
    """
    L_mot=[]
    for l in L:
        if(not process_word):
            l=processSentence(l)
        words= l.split(" ") #split each words
        for word in words:
            if(len(word)>=min_word_size): #dont keep small words
                L_mot.append(word)
    
    return collections.Counter(L_mot) #return dict {word:number of repeat}

def similarityBetween(target_account_description:str, follower_description:str, type_="hamming")->float:
    """
    Function to calcul similarity between two sentences (follower's description).
    This methode calcul similarity between two sentences taking into consideration the characters and not the meaning of the sentence

    :param str target_account_description: the first description (cleaning sentence)
    :param str follower_description: the second description (cleaning sentence)
    :param str type_: hamming or levenshtein

    :returns: similarity between two sentences
    :rtype: float
    """
    if(type_=="hamming"):
        return textdistance.hamming.similarity(target_account_description, follower_description) #distance #levenshtein #mlipns
    elif(type_=="levenshtein"):
        return textdistance.levenshtein.similarity(target_account_description, follower_description) #distance #levenshtein #mlipns

class autoReport:
    """Class to create an analyse docx after graph analyse : neuds distributed in class and size.

    Use only the constructor and the methode makeReport
    """
    def __init__(self, nodes_csv:str, edges_csv:str, classes_param:dict, account:dict, **kwargs):
        """
        :param str nodes_csv: path and nodes file
        :param str edges_csv: path and edges file
        :param dict classes_param: {class ID:color} color compatible with matplotlib : https://matplotlib.org/stable/gallery/color/named_colors.html
        :param dict account: {name:<str>, at:<str>, description:<str>}

        :param str desc_col_name: name of the description column nodes.csv (kwargs default "description")
        :param dict id_col_name: name of the id column nodes.csv (kwargs default "Id")
        :param str class_col_name: name of the class column nodes.csv (kwargs default "modularity_class")
        :param dict rank_col_name: name of the rank column nodes.csv (kwargs default "pageranks")
        :param str name_col_name: name of the name account column nodes.csv (kwargs default "name")
        :param dict at_col_name: name of the @ account column nodes.csv (kwargs default "screen_name")
        :param str source_col_name: name of the source column in edges.csv (kwargs default "Source")
        :param dict target_col_name: name of the target column in edges.csv (kwargs default "Target")
        """
        self.target_account=account
        self.classes=classes_param
        self.file_nodes_csv=nodes_csv
        self.file_edges_csv=edges_csv
        
        self._desc_col_name=kwargs.get("desc_col_name","description")
        self._id_col_name= kwargs.get("id_col_name","Id")
        self._class_col_name= kwargs.get("class_col_name","modularity_class")
        self._rank_col_name= kwargs.get("rank_col_name","pageranks")
        self._name_col_name= kwargs.get("name_col_name","name")
        self._at_col_name= kwargs.get("at_col_name","screen_name")
        
        self._source_col_name= kwargs.get("source_col_name","Source")
        self._target_col_name=kwargs.get("target_col_name","Target")
        
        self.node=pd.read_csv(self.file_nodes_csv).loc[:, [self._id_col_name, self._name_col_name,self._at_col_name, self._desc_col_name, self._class_col_name,self._rank_col_name]]
        self.edge=pd.read_csv(self.file_edges_csv).loc[:, [self._source_col_name, self._target_col_name]]
        self.nb_nodes= len(self.node)
        
    def class_distribution(self, file_name:str)->list:
        """
        Create a image file with a graph bar to show class distribution

        :param str file_name: name of the saving image (.png, .jpg...)
        :returns: list of distribution
        :rtype: list<float>
        """
        plt.clf()
        class_=collections.Counter(self.node.loc[:, self._class_col_name].tolist())
        
        plt.bar(list(self.classes.values()), [class_[c]/self.nb_nodes for c in self.classes.keys()], color=list(self.classes.values()))
        plt.xticks(rotation=45)
        plt.savefig(file_name, bbox_inches = 'tight')
        
        return [class_[c]/self.nb_nodes for c in self.classes.keys()] #return percentage 
    
    def most_common_words_tot(self, file_name:str, nb_words:int, min_len_word:int)->None:
        """
        Create a graph file with all most common words on followers description regardless of their class

        :param str file_name: name of the saving image (.png, .jpg...)
        :param int nb_words: number of the most common words to be displayed on the graph
        :param int min_len_word: minimum size of the words to be taken into consideration look :func:`most_commun_word`
        """
        plt.clf() #clear graph
        fig, ax = plt.subplots()

        countL_tot=most_commun_word(processListSentences(self.node.loc[:,self._desc_col_name].tolist()), min_len_word, True) #take all most commun word
        
        if("https" in countL_tot): #del https, it's not intresting
            countL_tot.pop("https")
        
        countL_tot=countL_tot.most_common(nb_words) #keep just some most common words
        
        #creation of the bar graph
        X = np.arange(len(countL_tot))*1.5
        for i,key in enumerate(self.classes.keys()):
            rslt_df = self.node[self.node[self._class_col_name] == key]
            L_phrase=processListSentences(rslt_df.loc[:,self._desc_col_name].tolist())
            countL=most_commun_word(L_phrase, min_len_word, True)

            ax.bar(X + (i/len(self.classes)-0.5), [countL[k]/val for k,val in countL_tot], color=self.classes[key], width = 1/len(self.classes))
        ax.set_xticks(X)
        ax.set_xticklabels([word for word, _ in countL_tot], rotation=45)
        ax.set_ylabel(f"distribution of the {len(countL_tot)} most common words by class")
        plt.savefig(file_name, bbox_inches = 'tight')
        plt.clf() #clear graph
    
    def most_common_words_class(self,file_name:str, key_class, nb_words:int, min_len_word:int):
        """
        Create a graph file with all most common words on followers description depending of their class

        :param str file_name: name of the saving image (.png, .jpg...)
        :param all key_class: key class : must be a key to classes_param
        :param int nb_words: number of the most common words to be displayed on the graph
        :param int min_len_word: minimum size of the words to be taken into consideration look :func:`most_commun_word`
        """
        plt.clf() #clear graph
        rslt_df = self.node[self.node[self._class_col_name] == key_class]
        L_phrase=processListSentences(rslt_df.loc[:,self._desc_col_name].tolist())
        countL=most_commun_word(L_phrase, min_len_word, True)
        if("https" in countL):
            countL.pop("https")
            
        L_=countL.most_common(nb_words)
        
        plt.bar([k for k, _ in L_], [val/len(L_phrase) for _, val in L_], width=1, color=self.classes[key_class]) #sum([val for _,val in L_])
        plt.xticks(rotation=45)
        plt.title(f"The {len(L_)} most common word in {self.classes[key_class]} class") #  {sum([val for k,val in class_ if k==key])/sum([val for _,val in class_])*100 :0.1f}%
        plt.savefig(file_name, bbox_inches = 'tight')
    
    def similarity_descriptions(self, file_name:str, similarity_fct):
        """
        Create a boxplot graph with similarity/distance between account target description and follower description; depending of their class

        :param str file_name: name of the saving image (.png, .jpg...)
        :param fct similarity_fct: function(str,str)->float, must return a similarity/distance 
        """
        plt.clf()
        result={}
        for key in self.classes.keys():
            rslt_df = self.node[self.node[self._class_col_name] == key]
            L_sentences=processListSentences(rslt_df.loc[:,self._desc_col_name].tolist())
            L_similarity=[]
            for sentence in L_sentences:
                L_similarity.append(similarity_fct(self.target_account["description"], sentence))

            result[key]=L_similarity #statistics.quantiles(data, *, n=4

        fig, ax = plt.subplots()
        ax.set_title(f"Similarities between the descriptions of the individuals in the class and {self.target_account['name']}'s description")
        ax.boxplot([result[key] for key in self.classes.keys()], showfliers=True, showmeans=True)
        ax.set_xticks(range(1,len(self.classes)+1))
        ax.set_xticklabels(list(self.classes.values()))
        ax.set_xlabel(f"class")
        plt.savefig(file_name, bbox_inches = 'tight')
    
    def links_between_class(self, file_name:str, percentage:float=1,node_weigth:list=None,logs:bool=True):
        """
        Create a node graph to resume the global graph, show all edges between classes
        This function can take time !

        :param str file_name: name of the saving image (.png, .jpg...)
        :param float percentage: must be in ]0,1], percentage of edge taken into consideration (in ordre of edges.csv) (default 1)
        :param list<float> node_weigth: same length as number of classes, weight of nodes, If None all nodes have the same weight (default None)
        :param bool logs: if plot some logs to show the progress (default True)
        """
        if(node_weigth is None):
            node_weigth=[1 for k in self.classes.keys()]
        
        plt.clf()
        id_={}
        A=[]
        i=0
        G=[]
        for i0,key_s in enumerate(self.classes.keys()):
            for j0,key_t in enumerate(self.classes.keys()):
                if(key_s != key_t):
                    A.append({"Source":key_s, "Target":key_t, "Weight":0})
                    id_[(key_s, key_t)]=i
                    i+=1
                    G.append((i0,j0))
        df_classe_edge=pd.DataFrame(A)
        
        N=int(len(self.edge)*min(percentage, 1))
        if(logs):
            print(f"links between classes with {percentage*100 : 0.1f}% of edges ({N} edges)")
        j=0
        for i in range(N):
            j+=1
            if(logs and j==(N//10)):
                j=0
                print(f"links between classes : {(i*100)/N : 0.1f}%")
            key_s=self.edge[self._source_col_name].iloc[i]
            key_t=self.edge[self._target_col_name].iloc[i]
            classe_s=self.node[self.node[self._id_col_name] == key_s][self._class_col_name].iloc[0]
            classe_t=self.node[self.node[self._id_col_name] == key_t][self._class_col_name].iloc[0]
            if((classe_s, classe_t) in id_.keys()):
                df_classe_edge.loc[id_[(classe_s, classe_t)], "Weight"]+=1
        
        g = igraph.Graph(directed=True)
        g.add_vertices(len(self.classes))

        for i, key in enumerate(self.classes.keys()):
            g.vs[i]["id"]= key
            g.vs[i]["label"]= "" #dictcolors[key][0]
            g.vs[i]["color"] = self.classes[key]

        H=[]
        W=[]
        seuil=1
        for i,g_ in enumerate(G):
            if(df_classe_edge.loc[i, "Weight"]>seuil):
                H.append(g_)
                W.append(df_classe_edge.loc[i, "Weight"])
        g.add_edges(H)
        g.es['width'] = [(w-min(W))/max(W)*15+3 for w in W]
        
        visual_style = {}
        # Set bbox and margin
        visual_style["bbox"] = (400,400)
        visual_style["margin"] = 27
        # Set vertex colours
        #visual_style["vertex_color"] = 'white'
        # Set vertex size
        visual_style["vertex_size"] = [50*l/max(node_weigth) for l in node_weigth]
        # Set vertex lable size
        visual_style["vertex_label_size"] = 22
        visual_style["edge_curved"] = [0.1 for i in range(len(H))]

        #igraph.plot(g, target=ax,**visual_style, edge_arrow_size= 1, directed = True)
        #plt.axis('off')
        #plt.savefig(file_name,bbox_inches = 'tight')
        out=igraph.plot(g,**visual_style, edge_arrow_size= 1, directed = True)
        out.save(file_name)

    def makeReport(self,file_name, logs=True, **kwargs):
        """
        Create a report : uses class methods

        :param str file_name: name of the report (.docx)
        :param bool logs: if plot some logs to show the progress (default True)
        :param int nb_words_tot: number of the most common words to be displayed on the graph (kwargs default 15)
        :param int nb_words_class: number of the most common words by class to be displayed on the graph (kwargs default 10)
        :param int min_len_word: minimum size of the words to be taken into consideration (kwargs default 5)
        :param int per_edeges: must be in ]0,1], percentage of edges taken into consideration (kwargs default 1)
        """
        file_class_distribution=f"class_distribution_{self.target_account['at']}.png"
        file_most_common_words_tot=f"most_common_words_tot_{self.target_account['at']}.png"
        file_most_common_words_class={key:f"most_common_words_{self.classes[key]}_{self.target_account['at']}.png" for key in self.classes.keys()}
        file_similarity_descriptions=f"similarity_descriptions_{self.target_account['at']}.png"
        file_links_between_class=f"links_between_class_{self.target_account['at']}.png"
        
        if(logs):
            print(f"starting {self.target_account['at']} report...")
            print(f"1st step : class distribution")
        
        pour_classe=self.class_distribution(file_class_distribution)
        
        nb_words=kwargs.get("nb_words_tot",15)
        min_len_word=kwargs.get("min_len_word",5)
        if(logs):
            print(f"2st step : most {nb_words} common words total with letters>={min_len_word}")
            
        self.most_common_words_tot(file_most_common_words_tot, nb_words=15, min_len_word=5)
        
        nb_words_class=kwargs.get("nb_words_class",10)
        
        if(logs):
            print(f"3st step : most {nb_words_class} common words by class with letters>={min_len_word}")
            
        for key in self.classes.keys():
            self.most_common_words_class(file_most_common_words_class[key],key_class=key, nb_words=nb_words_class, min_len_word=5)
        
        if(logs):
            print(f"4st step : similarity of descriptions")
            
        self.similarity_descriptions(file_similarity_descriptions, similarityBetween)
        
        if(logs):
            print(f"5st step : links between class /!\ It takes time !")
            
        self.links_between_class(file_links_between_class, percentage=kwargs.get("per_edeges",1), node_weigth=pour_classe, logs=logs)
        
        most_influential_accounts=[]
        for key in self.classes.keys():
            rslt_df = self.node[self.node[self._class_col_name] == key].sort_values(by=[self._rank_col_name], ascending=False)
            R=[]
            for i in range(min(len(rslt_df),10)):
                R.append({"name":rslt_df[self._name_col_name].iloc[i],
                          "at":"@"+rslt_df[self._at_col_name].iloc[i],
                          "description":rslt_df[self._desc_col_name].iloc[i]
                         })
            most_influential_accounts.append(R)
            
        self.writeDoc(file_name,target_account=self.target_account['name'], 
                target_account_at=self.target_account['at'],
                account_description=self.target_account['description'],
                number_classes=len(self.classes),
                file_class_distribution=file_class_distribution,
                file_most_commun_words_tot=file_most_common_words_tot,
                file_most_commun_words_classe=list(file_most_common_words_class.values()),
                most_commun_words_classe=10,
                most_commun_words_tot=15 , 
                      
                file_links_between_class=file_links_between_class,
                file_similarity_descriptions=file_similarity_descriptions,
                color_classe=list(self.classes.values()),
                pour_classe=[p*100 for p in pour_classe],

                most_influential_accounts=most_influential_accounts #[[{"name":"jp", "at":"@jp", "description":"bla"},{"name":"jp", "at":"@jp", "description":"bla"}]]
                )
    
    def writeDoc(self,file_name, **kwargs):
        """
        write a report in docx

        :param str file_name: name of the doc (.docx)
        :param str target_account: name of the target account (kwargs default Nan)
        :param str target_account_at: at@ of the target account (kwargs default @Nan)
        :param str account_description: account description (kwargs default "")

        :param str number_classes: number of all classes (kwargs default )
        :param list<str> color_classe: color classe (kwargs default [])
        :param list<float> pour_classe: pourcentage of nodes in classe (kwargs default [])

        :param str file_class_distribution: file where is the graph class distribution graph (kwargs default "")
        :param str file_most_commun_words_tot: file where is the graph class most commun words total (kwargs default "")
        :param list<str> str file_most_commun_words_classe: files where is graphs most commun words by class (kwargs default "")
        :param str file_similarity_descriptions: file where is the graph similarity descriptions (kwargs default "")
        :param str file_links_between_class: file where is the graph links between class (kwargs default "")
        
        :param int most_commun_words_tot: number of most commun words total (kwargs default )
        :param int most_commun_words_classe: number of most commun words by class (kwargs default )

        :param list<list<dict>> most_influential_accounts: [[{"name":"", "at":"@", "description":""}]] (kwargs default )
        """
        target_account=kwargs.get("target_account", "Nan")
        target_account_at=kwargs.get("target_account_at", "@Nan")
        number_classes=kwargs.get("number_classes", 1)
        file_class_distribution=kwargs.get("file_class_distribution", "graph.png")
        
        most_commun_words_tot=kwargs.get("most_commun_words_tot", 10)
        file_most_commun_words_tot=kwargs.get("file_most_commun_words_tot", "graph.png")
        file_links_between_class=kwargs.get("file_links_between_class", "graph.png")
        file_similarity_descriptions=kwargs.get("file_similarity_descriptions", "graph.png")
        account_description=kwargs.get("account_description", "super account_description")
        
        most_commun_words_classe=kwargs.get("most_commun_words_classe", 10)
        file_most_commun_words_classe=kwargs.get("file_most_commun_words_classe", ["graph.png"])

        color_classe=kwargs.get("color_classe", ["red"])
        pour_classe=kwargs.get("pour_classe", [10])

        most_influential_accounts=kwargs.get("most_influential_accounts", [[{"name":"jp", "at":"@jp", "description":"bla"},{"name":"jp", "at":"@jp", "description":"bla"}]])


        document = Document()

        document.add_heading(f'Analysis of {target_account} ({target_account_at})\'s Graph', 0)

        document.add_heading('Introduction', level=1)

        p = document.add_paragraph('')
        p.add_run(f'<insert a comment about {target_account}>').italic = True

        p = document.add_paragraph("""On the document below, we will first see a global analysis, and a comparative analysis of the class, then we will see in each class the most influential accounts""")


        ##_______

        document.add_heading('Class distribution', level=2)
        p = document.add_paragraph(f'All individuals are placed into one of {number_classes} class :')

        document.add_picture(file_class_distribution, width=Inches(5))

        ##_______

        document.add_heading('Links between class', level=2)
        p = document.add_paragraph(f"The most interesting analysis for the comparison between the class, is to show links between class.")
        p = document.add_paragraph(f"")
        p.add_run(f'All accounts of a class are merged to a sigle node and we see that the links that leave the class').italic = True

        document.add_picture(file_links_between_class, width=Inches(5))

        p = document.add_paragraph(f"")
        p.add_run(f'<insert a comment about this graph>').italic = True

        ##_______

        document.add_heading('Most common words', level=2)
        p = document.add_paragraph(f"For an initial understanding, the target's graph, we can see the {most_commun_words_tot} most commun words on follower's descriptions spread on each classe")

        document.add_picture(file_most_commun_words_tot, width=Inches(5))

        p = document.add_paragraph(f"")
        p.add_run(f'<insert a comment about this graph>').italic = True

        ##_______

        document.add_heading('Similarity of descriptions', level=2)
        p = document.add_paragraph(f"Similarity between the follower's description and the description of {target_account}.")
        document.add_paragraph(account_description, style='Intense Quote')

        document.add_picture(file_similarity_descriptions, width=Inches(5))

        document.add_paragraph(f"This methode use only the similarity between {target_account}'s description and follower's description, and not the meaning of the description")

        p = document.add_paragraph(f"")
        p.add_run(f'<insert a comment about this graph>').italic = True

        document.add_heading('Class analysis', level=1)
        for i in range(number_classes):
            document.add_heading(f'Class {color_classe[i]} ({pour_classe[i] : .1f}%)', level=2)

            document.add_heading(f'Most {most_commun_words_classe} common words', level=3)
            document.add_paragraph(f"We only check for words with more than 5 characters")
            document.add_picture(file_most_commun_words_classe[i], width=Inches(5))

            document.add_heading(f'Most influential accounts', level=3)
            #p=document.add_paragraph("", style='List Number')
            for k in range(len(most_influential_accounts[i])):
                p=document.add_paragraph(f"{most_influential_accounts[i][k]['name']} ({most_influential_accounts[i][k]['at']}) : «{most_influential_accounts[i][k]['description']}» ", style='List Bullet')

        document.save(file_name)

if __name__ == '__main__':
    r=autoReport("nodes.csv", "edges.csv", {2339:"magenta", 2433:"limegreen", 2566:"deepskyblue", 1052:"saddlebrown", 2498:"darkorange", 1583:"red"}, 
           {"name":"Antoine Bondaz", "at":"@AntoineBondaz", "description":"""Foodie - 🕵🏼‍Research @FRS_org - 👨🏼‍🏫 Teach @SciencesPo - 🇨🇳🇹🇼🇰🇷🇰🇵's foreign & security policies - Ph.D."""})

    r.makeReport("out.docx")