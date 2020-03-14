#!/usr/bin/env python
# coding: utf-8
# D:\python project me\TG下多txt输入，多txt输出
import nltk
# nltk.download('stopwords')
from nltk.corpus import stopwords
from nltk.cluster.util import cosine_distance
import numpy as np
import networkx as nx
import os
import codecs
import comtypes
from comtypes.client import CreateObject
class generatesumm:

    def read_article(self,file_name):
        file = open(file_name,"r",encoding='UTF-8')
        filedata = file.readlines()
       # article = filedata[0].split(str.encode(','))
        article = filedata[0].split(",")
        sentences = []

        for sentence in article:
          #  print(sentence)
            #sentences.append(sentence.replace(str.encode('[^a-zA-Z]'), str.encode(' ')).split(str.encode(' ')))
            sentences.append(sentence.replace("[^a-zA-Z]", " ").split(" "))
        sentences.pop()

        return sentences

    def sentence_similarity(self, sent1, sent2, stopwords=None):
        if stopwords is None:
           stopwords = []

        sent1 = [w.lower() for w in sent1]
        sent2 = [w.lower() for w in sent2]

        all_words = list(set(sent1 + sent2))

        vector1 = [0] * len(all_words)
        vector2 = [0] * len(all_words)

    # build the vector for the first sentence
        for w in sent1:
           if w in stopwords:
              continue
        vector1[all_words.index(w)] += 1

    # build the vector for the second sentence
        for w in sent2:
            if w in stopwords:
              continue
        vector2[all_words.index(w)] += 1

        return 1 - cosine_distance(vector1, vector2)


    def build_similarity_matrix(self, sentences, stop_words):
    # Create an empty similarity matrix
       similarity_matrix = np.zeros((len(sentences), len(sentences)))

       for idx1 in range(len(sentences)):
           for idx2 in range(len(sentences)):
               if idx1 == idx2:  # ignore if both are same sentences
                  continue
               similarity_matrix[idx1][idx2] = self.sentence_similarity(sentences[idx1], sentences[idx2], stop_words)

       return similarity_matrix


    def generate_summary(self, file_name, top_n=5):
        nltk.download("stopwords")
        stop_words = stopwords.words('english')
        summarize_text = []

    # Step 1 - Read text anc split it
        sentences = self.read_article(file_name)

    # Step 2 - Generate Similary Martix across sentences
        sentence_similarity_martix = self.build_similarity_matrix(sentences, stop_words)

    # Step 3 - Rank sentences in similarity martix
        sentence_similarity_graph = nx.from_numpy_array(sentence_similarity_martix)
        scores = nx.pagerank(sentence_similarity_graph)

    # Step 4 - Sort the rank and pick top sentences
        ranked_sentence = sorted(((scores[i], s) for i, s in enumerate(sentences)), reverse=True)
       # print("Indexes of top ranked_sentence order are ", ranked_sentence)

        for i in range(top_n):
            summarize_text.append(" ".join(ranked_sentence[i][1]))
        summarize_texted=". ".join(summarize_text)
        return summarize_texted
    # Step 5 - Offcourse, output the summarize texr
        #print("Summarize Text: \n", ". ".join(summarize_text))
        #print("lllaaaaaa\n",summarize_text)

    # let's begin

class readfile:
    def init_txt(self):
        txt = comtypes.client.CreateObject("txt.Application","doc.Application","docx.Application","xml.Application")
        txt.Visible = 1
        return txt

    def readthefile(self, folder):
        # 获取指定目录下面的所有文件
        files = os.listdir(folder)
        # 获取word类型的文件放到一个列表里面
        wdfiles = [f for f in files if f.endswith((".doc", ".docx","txt","xml"))]
        conclusion = []  #文摘的结果
        for wdfile in wdfiles:
         # 将word文件放到指定的路径下面
            # wdPath = os.path.join("D:\python project me\TG", wdfile)
            filethename=wdfile
            objj=generatesumm()
            #创建方法
            ww=objj.generate_summary(filethename,3)
            conclusion.append(ww)
        #    print(conclusion)
            for theconclusion in conclusion:
                #将文摘句子放到新建txt里
                outfilethename="summary"+filethename
                #新建txt命名为"summary+源文件名字"
                file = open(outfilethename, 'w')
                file.write(str(theconclusion));
                file.close()

     #   print("conclusion\n",conclusion)
      #  print("conclusion\n",".".join(conclusion))

if __name__ == '__main__':
    #obj = generatesumm()
    obj= readfile()
  #  obj.readthefile("D:\python project me\TG\projectartical")
    obj.readthefile('D:\python project me\TG')
   # obj.generate_summary("msft.txt", 2)
