import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from sklearn.decomposition import LatentDirichletAllocation
import umap
import matplotlib.pyplot as plt
import os
from pptx import Presentation
from pptx.util import Pt, Inches
from transformers import GPT2LMHeadModel, GPT2Tokenizer
import datetime