from docx import Document
from docx.shared import Inches

import numpy as np
import pandas as pd
import collections
import matplotlib.pyplot as plt
import textdistance
import statistics 
import igraph

def processWord(word):
    caraToDel="🐝🚲⛺🌻🌲⭐⚽👨🏼‍🏫🕵🧭🇯🇻🏼‍😃👨🏻‍🏫🌎📲🏛🗽🌊🚢"  #del som emoji
    word=word.lower()
    transTable = word.maketrans("éèà-'.ôç≠ę:î", "eea   oc e i", "!#$%^&*()\"’‘«»,，@_/|"+caraToDel)
    word = word.translate(transTable)
    #translate some emoji
    for drap, txt in zip(["🇫🇷", "🇨🇳", "🇹🇼","🇰🇷","🇰🇵","🇩🇪","🇺🇲","🇪🇺", "🇷🇺", "🇮🇷","🇺🇳"], ["france","china", "taiwan", "south korea", "north korea", "germany", "usa", "europe", "russia", "iran","un"]):
        if(drap in word):
            word=word[:word.index(drap)]+" "+txt+" "+word[word.index(drap)+2:]
    while "\n" in word:
        word=word[:word.index("\n")]+" "+word[word.index("\n")+1:]
    while "  " in word:
        word=word[:word.index("  ")]+" "+word[word.index("  ")+2:]
    return word

def most_commun_word(L:"list<str>", min_word_size=3, process_word=False):
    L_mot=[]
    for l in L:
        if(not process_word):
            l=processWord(l)
        words= l.split(" ")
        for word in words:
            if(len(word)>=min_word_size):
                L_mot.append(word)
    
    return collections.Counter(L_mot)

def process_phrase(L:"list<str>"):
    L_phrase=[]
    for l in L:
        a=processWord(str(l))
        L_phrase.append(a)
    return L_phrase

def similarityBetween(target_account_description:str, follower_description:str, type_="hamming")->float:
    if(type_=="hamming"):
        return textdistance.hamming.similarity(target_account_description, follower_description) #distance #levenshtein #mlipns
    elif(type_=="levenshtein"):
        return textdistance.levenshtein.similarity(target_account_description, follower_description) #distance #levenshtein #mlipns

class autoReport:
    def __init__(self, node_csv:str, edge_csv:str, classes_param:dict, account:dict, **kwargs):
        self.target_account=account # keys: name, at, description
        self.classes=classes_param #keys : id_classes:color color compatible with matplotlib : https://matplotlib.org/stable/gallery/color/named_colors.html
        self.file_node_csv=node_csv
        self.file_edge_csv=edge_csv
        
        self._desc_col_name=kwargs.get("desc_col_name","description")
        self._id_col_name= kwargs.get("id_col_name","Id")
        self._class_col_name= kwargs.get("class_col_name","modularity_class")
        self._rank_col_name= kwargs.get("rank_col_name","pageranks")
        self._name_col_name= kwargs.get("name_col_name","name")
        self._at_col_name= kwargs.get("at_col_name","screen_name")
        
        self._source_col_name= kwargs.get("source_col_name","Source")
        self._target_col_name=kwargs.get("target_col_name","Target" )
        
        self.node=pd.read_csv(node_csv).loc[:, [self._id_col_name, self._name_col_name,self._at_col_name, self._desc_col_name, self._class_col_name,self._rank_col_name]]
        self.edge=pd.read_csv(edge_csv).loc[:, [self._source_col_name, self._target_col_name]]
        self.nb_nodes= len(self.node)
        
    def class_distribution(self, file_name):
        plt.clf()
        class_=collections.Counter(self.node.loc[:, self._class_col_name].tolist())
        
        plt.bar(list(self.classes.values()), [class_[c]/self.nb_nodes for c in self.classes.keys()], color=list(self.classes.values()))
        plt.xticks(rotation=45)
        plt.savefig(file_name, bbox_inches = 'tight')
        
        return [class_[c]/self.nb_nodes for c in self.classes.keys()]
    
    def most_common_words_tot(self, file_name, nb_words, min_len_word):
        plt.clf()
        fig, ax = plt.subplots()

        countL_tot=most_commun_word(process_phrase(self.node.loc[:,self._desc_col_name].tolist()), min_len_word, True)
        
        if("https" in countL_tot):
            countL_tot.pop("https")
        
        countL_tot=countL_tot.most_common(nb_words)
        
        X = np.arange(len(countL_tot))*1.5

        for i,key in enumerate(self.classes.keys()):
            rslt_df = self.node[self.node[self._class_col_name] == key]
            L_phrase=process_phrase(rslt_df.loc[:,self._desc_col_name].tolist())
            countL=most_commun_word(L_phrase, min_len_word, True)

            ax.bar(X + (i/len(self.classes)-0.5), [countL[k]/val for k,val in countL_tot], color=self.classes[key], width = 1/len(self.classes))
        ax.set_xticks(X)
        ax.set_xticklabels([word for word, _ in countL_tot], rotation=45)
        ax.set_ylabel(f"distribution of the {len(countL_tot)} most common words by class")
        plt.savefig(file_name, bbox_inches = 'tight')
        plt.clf()
    
    def most_common_words_class(self,file_name, key_class, nb_words, min_len_word):
        plt.clf()
        rslt_df = self.node[self.node[self._class_col_name] == key_class]
        L_phrase=process_phrase(rslt_df.loc[:,self._desc_col_name].tolist())
        countL=most_commun_word(L_phrase, min_len_word, True)
        if("https" in countL):
            countL.pop("https")
            
        L_=countL.most_common(nb_words)
        
        plt.bar([k for k, _ in L_], [val/len(L_phrase) for _, val in L_], width=1, color=self.classes[key_class]) #sum([val for _,val in L_])
        plt.xticks(rotation=45)
        plt.title(f"The {len(L_)} most common word in {self.classes[key_class]} class") #  {sum([val for k,val in class_ if k==key])/sum([val for _,val in class_])*100 :0.1f}%
        plt.savefig(file_name, bbox_inches = 'tight')
    
    def similarity_descriptions(self, file_name, similarity_fct):
        plt.clf()
        result={}
        for key in self.classes.keys():
            rslt_df = self.node[self.node[self._class_col_name] == key]
            L_sentences=process_phrase(rslt_df.loc[:,self._desc_col_name].tolist())
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
        g.es['width'] = [w/max(W)*15 for w in W]
        
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
        visual_style["edge_curved"] = 0.1

        igraph.plot(g, file_name, **visual_style, edge_arrow_size= 1, directed = True)
    
    def makeReport(self,file_name, logs=True, **kwargs):
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
            for i in range(10):
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