#!/usr/bin/env python
# coding: utf-8
#
#中文词向量
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
#from word_embeddings_file import word_embeddings
from rouge import Rouge
from itertools import chain
import gensim

from gensim.models import KeyedVectors
import numpy as np
import re,os,jieba

class generatesumm:

    def read_article(self,file_name,folder):
        sentences = []  #存放去掉空行和\n的句子
        with open(folder+"\\"+file_name, 'r',encoding='UTF-8') as f1:
            for line in f1.readlines():
               # if line != None:
               #     sentences.append(line.strip("\n"))   #去掉句子结尾的\n
                    if line.strip():
                        # 把元素按照[。！；？]进行分隔，得到句子。
                        line_split = re.split(r'[。！；？]', line.strip())
                        # [。！；？]这些符号也会划分出来，把它们去掉。
                        line_split = [line.strip() for line in line_split if
                                      line.strip() not in ['。', '！', '？', '；'] and len(line.strip()) > 1]
                        sentences.append(line_split)
            sentences_list= list(chain.from_iterable(sentences))
        sentences_lists=[]
        f1.close()
        for sentence in sentences_list:
            sentences_de=self.de_nonchinese(sentence)
            sentences_lists.append(sentences_de)
        sentences_lists = [i for i in sentences_lists if (len(str(i)) != 0)]    #去掉空的句子
        #sentences = re.sub(r'[^\u4e00-\u9fa5]+', '', sentences)  #去掉非汉字

        sentence_word_list = []
        for sentence in sentences_lists:
            line_seg = self.seg_depart(sentence)
            sentence_word_list.append(line_seg)
       # print((sentence_word_list))
        return sentence_word_list,sentences_lists
    def de_nonchinese(self,sentence):
        # 去掉非汉字字符
        sentence = re.sub(r'[^\u4e00-\u9fa5]+','',sentence)
        return  sentence

    def seg_depart(self,sentence):
        # 分词
       # sentence = re.sub(r'[^\u4e00-\u9fa5]+','',sentence)

        sentence_depart = [word for word in jieba.cut(sentence.strip())]
        return sentence_depart



    def remove_stopwords(self,sen,stop_words=None):
        if stop_words is None:
            stop_words=[]
        else:
            sen_new = " ".join([i for i in sen if i not in stop_words])
        return sen_new

    def sentence_vectors(self,sentences, word_embeddings, stopwords=None ):
        if stopwords is None:
           stopwords = []
       # clean_sentences=[self.remove_stopwords(r.split(),stopwords) for r in sentences]
        clean_sentences = [self.remove_stopwords(r,stopwords) for r in sentences]


        sentence_vectors = []    #句子向量
        for i in clean_sentences:
            if len(i) != 0:
                #v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i.split()]) / (len(i.split()) + 0.001)
               # v = sum([word_embeddings.get(w, np.zeros((1,))) for w in i]) / (len(i) + 0.001)
                v = sum([word_embeddings.get(w, np.random.uniform(0, 1, 300)) for w in i]) / (len(i) + 0.001)
            else:
              #  v = np.zeros((30,))
                v = np.random.uniform(0, 1, 300)
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

    def generate_summary(self, file_name, folder,stop_words,word_embeddings,top_n=5):

        summarize_text = []

    # Step 1 - Read text anc split it
        sentences ,filedataa= self.read_article(file_name,folder)
        sentence_vectors=self.sentence_vectors(sentences,word_embeddings,stop_words)
    # Step 2 - Generate Similary Martix across sentences
        sentence_similarity_martix = self.build_similarity_matrix(sentences, sentence_vectors,stop_words)

    # Step 3 - Rank sentences in similarity martix
        sentence_similarity_graph = nx.from_numpy_array(sentence_similarity_martix)
        scores = nx.pagerank(sentence_similarity_graph)

    # Step 4 - Sort the rank and pick top sentences
        ranked_sentence = sorted(((scores[i], s) for i, s in enumerate(filedataa)), reverse=True)
       # print("Indexes of top ranked_sentence order are ", ranked_sentence)

        for i in range(top_n):
            summarize_text.append("".join(ranked_sentence[i][1]))
        summarize_texted="。 ".join(summarize_text)
        return summarize_texted,filedataa
    # Step 5 - Offcourse, output the summarize texr
        #print("Summarize Text: \n", ". ".join(summarize_text))
        #print("lllaaaaaa\n",summarize_text)

    # let's begin

class readfile:
    def init_txt(self):
        txt = comtypes.client.CreateObject("txt.Application","doc.Application","docx.Application","xml.Application")
        txt.Visible = 1
        return txt

    def readthefile(self, folder,word_embeddings):
        # 获取指定目录下面的所有文件
        #nltk.download("stopwords")
        #stop_words = stopwords.words('english')
       # stop_words = stopwords.words('zh')
        stop_words = [line.strip() for line in open('./stopwords.txt', encoding='UTF-8').readlines()]
        files = os.listdir(folder)
        # 获取word类型的文件放到一个列表里面
        wdfiles = [f for f in files if f.endswith((".doc", ".docx","txt","xml"))]

        n=1
        rouge_resulting = []
        for wdfile in wdfiles:
         # 将word文件放到指定的路径下面
            # wdPath = os.path.join("D:\python project me\TG", wdfile)
            conclusion=[]  #文摘的结果
            filethename=wdfile
            objj=generatesumm()
            #创建方法
            ww,line_split=objj.generate_summary(filethename,folder,stop_words,word_embeddings,3)
            conclusion.append(ww)

            a = conclusion  # 预测摘要 （可以是列表也可以是句子）
            c= line_split[0]
            b = [c]  # 真实摘要
            #print(a)
           # print(b)

            print(n)
            n+=1
            print(filethename)
            '''
            f:F1值  p：查准率  R：召回率
            '''
            rouge = Rouge()
            rouge_score = rouge.get_scores(a, b)
           # print(filethename+" "+"rouge:")
           # print(rouge_score[0]["rouge-1"])
            rouge_1="rouge-1:"+str(rouge_score[0]["rouge-1"])
           # print(rouge_score[0]["rouge-2"])
            rouge_2="rouge-2:"+str(rouge_score[0]["rouge-2"])
            #print(rouge_score[0]["rouge-l"])
            rouge_3="rouge-L"+str(rouge_score[0]["rouge-l"])

            rouge_resulting.append(filethename)
            rouge_resulting.append(rouge_1)
            rouge_resulting.append(rouge_2)
            rouge_resulting.append(rouge_3)

            for theconclusion in conclusion:
                #将文摘句子放到新建txt里
                outfilethename="summary"+filethename
                #新建txt命名为"summary+源文件名字"
                file = open(folder+"\\"+"out"+"\\"+outfilethename, 'w', encoding='utf-8')
                file.write(str(theconclusion));
                file.close()

        rougefile = open(folder + "\\" + "out" + "\\" + "rouge_result.txt", 'w', encoding='utf-8')
        for i in rouge_resulting:
            s = i + "\n"
            rougefile.write(s)
        file.close()



if __name__ == '__main__':
    word_embeddings = {}
    f = open('D:\python project me\sgns.financial.char\sgns.financial.char', encoding='utf-8')
    for line in f:
        # 把第一行的内容去掉
        if '467389 300\n' not in line:
            values = line.split()
            # 第一个元素是词语
            word = values[0]
            embedding = np.asarray(values[1:], dtype='float32')
            word_embeddings[word] = embedding
    f.close()
    print("一共有" + str(len(word_embeddings)) + "个词语/字。")
    #word_embeddings=word_embeddings()
    print("www")
    obj= readfile()
    obj.readthefile('D:\python project me\mtext\koq',word_embeddings)
    #输出是D:\python project me\mtext\mbusiness\out文件夹下

