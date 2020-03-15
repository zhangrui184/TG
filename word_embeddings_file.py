# coding: utf-8
# 没有用到
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


class word_embeddings:


    def word_embeddings(self):
        #if word_embeddings is None:
        word_embeddings = {}  # 词向量
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
        return word_embeddings