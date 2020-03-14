#!/usr/bin/env python
# coding: utf-8
# nltk传stop_words改到gennery summary()
#自己写的sentence_similarity方法
#错误是reshape映射问题
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
#from sklearn.metrics.pairwise import cosine_similarity
class generatesumm:

    def read_article(self,file_name,folder):
        file = open(folder+"\\"+file_name,"r",encoding='UTF-8')
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

    def remove_stopwords(self,sen,stop_words=None):
        if stop_words is None:
            stop_words=[]
        else:
            sen_new = " ".join([i for i in sen if i not in stop_words])
        return sen_new

    def sentence_vectors(self,sentences, stopwords=None):
        if stopwords is None:
           stopwords = []
       # clean_sentences=[self.remove_stopwords(r.split(),stopwords) for r in sentences]
        clean_sentences = [self.remove_stopwords(r,stopwords) for r in sentences]
        word_embeddings = {}   #词向量
        f = open('glove.42B.300d.txt', encoding='utf-8')
        for line in f:
            values = line.split()
            word = values[0]
            coefs = np.asanyarray(values[1:], dtype='float32')
            word_embeddings[word] = coefs
        f.close()

        sentence_vectors = []    #句子向量
        for i in sentences:
            if len(i) != 0:
                #v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i.split()]) / (len(i.split()) + 0.001)
                v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i]) / (len(i) + 0.001)
            else:
                v = np.zeros((30,))
            sentence_vectors.append(v)

        return sentence_vectors


    def sentence_similarity(self, sent1, sent2):
       # if stopwords is None:
        #   stopwords = []
         sent1w=sent1
         sent2w=sent2

         all_words = list(set(sent1w + sent2w))

         score=[0] * len(all_words)
         score=1 - cosine_distance(sent1w, sent2w)
     #   all_words = list(set(sent1 + sent2))

       # vector1 = self.sentence_vectors[0]
     #   vector2 = self.sentence_vectors[0]
         # return 1 - cosine_distance(vector1, vector2)
         return score

    def build_similarity_matrix(self, sentences, sentence_vectors, stop_words):
    # Create an empty similarity matrix
       similarity_matrix = np.zeros((len(sentences), len(sentences)))

       for idx1 in range(len(sentences)):
           for idx2 in range(len(sentences)):
               if idx1 == idx2:  # ignore if both are same sentences
                  continue
              # similarity_matrix[idx1][idx2] = self.sentence_similarity(sentence_vectors[idx1].reshape(1,1),sentence_vectors[idx2].reshape(1,1))
              # similarity_matrix[idx1][idx2] = self.sentence_similarity(sentence_vectors[idx1],sentence_vectors[idx2])
               similarity_matrix[idx1][idx2] = self.sentence_similarity(sentence_vectors[idx1].reshape(1, 1),sentence_vectors[idx2].reshape(1, 1))
       return similarity_matrix

    def generate_summary(self, file_name, folder,stop_words,top_n=5):

        summarize_text = []

    # Step 1 - Read text anc split it
        sentences = self.read_article(file_name,folder)
        sentence_vectors=self.sentence_vectors(sentences,stop_words)
    # Step 2 - Generate Similary Martix across sentences
        sentence_similarity_martix = self.build_similarity_matrix(sentences, sentence_vectors,stop_words)

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
        nltk.download("stopwords")
        stop_words = stopwords.words('english')
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
            ww=objj.generate_summary(filethename,folder,stop_words,3)
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
    #obj.readthefile('D:\python project me\TG')
    obj.readthefile('D:\python project me\TG\kos')
   # obj.generate_summary("msft.txt", 2)
