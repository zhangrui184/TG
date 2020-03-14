#!/usr/bin/env python
# coding: utf-8
#
#改读文件后sentences的表达方式，去掉空行和\n
#成功将txt转到out文件夹下的summary+name+txt
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
from sklearn.metrics.pairwise import cosine_similarity
class generatesumm:

    def read_article(self,file_name,folder):
        sentences = []  #存放去掉空行和\n的句子
        with open(folder+"\\"+file_name, 'r',encoding='UTF-8') as f1:
            for ip in f1.readlines():
                if ip != None:
                    sentences.append(ip.strip("\n"))   #去掉句子结尾的\n
        f1.close()
        sentences = [i for i in sentences if (len(str(i)) != 0)]    #去掉空的句子

        filedataa = []   #存放每个用逗号隔开的句子
        for i in range(len(sentences)):
            filedataa += sentences[i].split(",")
      #  filedataa.remove('\n')
        sentencess = [] #存放统一小写的句子

        for sentence in filedataa:
          #  print(sentence)
            #sentences.append(sentence.replace(str.encode('[^a-zA-Z]'), str.encode(' ')).split(str.encode(' ')))
            sentencess.append(sentence.replace("[^a-zA-Z]", " ").split(" "))
        sentencess.pop()

        return sentencess

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
            try:
                word_embeddings[word] = coefs
            except Exception:
                word_embeddings[word] = np.random.uniform(0, 1, 300)
        f.close()

        sentence_vectors = []    #句子向量
        for i in sentences:
            if len(i) != 0:
                #v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i.split()]) / (len(i.split()) + 0.001)
               # v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i]) / (len(i) + 0.001)
                v = sum([word_embeddings.get(w, np.random.uniform(0, 1, 300)) for w in i]) / (len(i) + 0.001)
            else:
              #  v = np.zeros((30,))
                v = np.zeros((1,))
            sentence_vectors.append(v)

        return sentence_vectors

    def build_similarity_matrix(self, sentences, sentence_vectors, stop_words):
    # Create an empty similarity matrix
       similarity_matrix = np.zeros((len(sentences), len(sentences)))

       for idx1 in range(len(sentences)):
           for idx2 in range(len(sentences)):
              # if idx1 == idx2:  # ignore if both are same sentences
               if idx1 != idx2:  # ignore if both are same sentences
                  #continue
              # similarity_matrix[idx1][idx2] = self.sentence_similarity(sentence_vectors[idx1].reshape(1,1),sentence_vectors[idx2].reshape(1,1))
              # similarity_matrix[idx1][idx2] = self.sentence_similarity(sentence_vectors[idx1],sentence_vectors[idx2])
                 similarity_matrix[idx1][idx2] = cosine_similarity(sentence_vectors[idx1].reshape(1,300),sentence_vectors[idx2].reshape(1,300))[0,0]
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
                file = open(folder+"\\"+"out"+"\\"+outfilethename, 'w')
                file.write(str(theconclusion));
                file.close()

     #   print("conclusion\n",conclusion)
      #  print("conclusion\n",".".join(conclusion))

if __name__ == '__main__':
    #obj = generatesumm()
    obj= readfile()
  #  obj.readthefile("D:\python project me\TG\projectartical")
    #obj.readthefile('D:\python project me\TG')
    obj.readthefile('D:\python project me\TG\kos\mskj')
   # obj.generate_summary("msft.txt", 2)
