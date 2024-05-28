# import weaviate
# from langchain.embeddings import AzureOpenAIEmbeddings
# import sys, os
# sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# import utils.config_log as config_log
# config, logger, CONFIG_PATH = config_log.setup_config_and_logging()
# config.read(CONFIG_PATH)

# Azure_Open_AI_VERSION = 'Open_AI'
# os.environ["OPENAI_API_TYPE"] = config.get(Azure_Open_AI_VERSION, 'api_type')
# os.environ["OPENAI_API_VERSION"] = config.get(Azure_Open_AI_VERSION, 'api_version')
# os.environ["OPENAI_API_BASE"] = config.get(Azure_Open_AI_VERSION, 'azure_endpoint')
# os.environ["OPENAI_API_KEY"] = config.get(Azure_Open_AI_VERSION, 'api_key')
# wea_url = config.get('Weaviate', 'weaviate_url')
# api_key = config.get(Azure_Open_AI_VERSION, 'api_key')
# vdb = config.get('Weaviate', 'classnm')

# class WeaviateSemanticSearch:
#     def __init__(self, classNm):
#         self.url = wea_url
#         self.embeddings = AzureOpenAIEmbeddings(chunk_size=1, model="text-embedding-ada-002", api_key=api_key)
#         self.client = weaviate.Client(
#             url=wea_url,
#             additional_headers={"X-Azure-Api-Key": api_key}
#         )
#         self.classNm = classNm

#     def hybrid_search(self, query_text, num, alpha):
#         query_vector = self.embeddings.embed_query(query_text)

#         vector_str = ",".join(map(str, query_vector))

#         gql_query = f"""
#         {{
#             Get {{
#                 {self.classNm}(hybrid: {{query: "{query_text}", vector: [{vector_str}], alpha: {alpha} }}, limit: {num}) {{
#                     content
#                     _additional {{
#                         distance
#                         score
#                     }}
#                 }}
#             }}
#         }}
#         """

#         search_results = self.client.query.raw(gql_query)
        
#         if 'errors' in search_results:
#             raise Exception(search_results['errors'][0]['message'])
        
#         results = search_results['data']['Get'][self.classNm]
        
#         return results

# def search_do(input_, alp):
#     searcher = WeaviateSemanticSearch(vdb)
#     results = searcher.hybrid_search(input_, 3, alpha=alp)

#     result_li = []
#     for _, result in enumerate(results, 1):
#         result_li.append(result['content'])

#     return result_li

# if __name__ == "__main__":
#     # An alpha of 1 is a pure vector search
#     # An alpha of 0 is a pure keyword search
#     quest = "老師 蔡炎龍 學生"
#     with open("result.txt", 'w', encoding='utf-8') as file:
#         file.write(str(search_do(quest, alp=1)) + "\n\n\n" + str(search_do(quest, alp=0.1)))

import weaviate
from sentence_transformers import SentenceTransformer
import sys, os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import utils.config_log as config_log
config, logger, CONFIG_PATH = config_log.setup_config_and_logging()
config.read(CONFIG_PATH)

wea_url = config.get('Weaviate', 'weaviate_url')
vdb = config.get('Weaviate', 'classnm')
PROPERTIES = ["uuid", "title", "content"]

class WeaviateSemanticSearch:
    def __init__(self, classNm):
        self.url = wea_url
        self.embeddings = SentenceTransformer('all-MiniLM-L6-v2')
        self.client = weaviate.Client(url=wea_url)
        self.classNm = classNm

    def aggregate_count(self):
        return self.client.query.aggregate(self.classNm).with_meta_count().do()

    def get_all_data(self, limit=3):
        if self.client.schema.exists(self.classNm):
            result = self.client.query.get(class_name=self.classNm, properties=PROPERTIES).with_limit(limit).do()
            return result
        else:
            raise Exception(f"Class {self.classNm} does not exist.")

    def delete_class(self):
        self.client.schema.delete_class(self.classNm)

    def hybrid_search(self, query_text, num, alpha):
        query_vector = self.embeddings.encode(query_text)
        vector_str = ",".join(map(str, query_vector))
        gql_query = f"""
        {{
            Get {{
                {self.classNm}(hybrid: {{query: "{query_text}", vector: [{vector_str}], alpha: {alpha} }}, limit: {num}) {{
                    content
                    _additional {{
                        distance
                        score
                    }}
                }}
            }}
        }}
        """
        search_results = self.client.query.raw(gql_query)

        if 'errors' in search_results:
            raise Exception(search_results['errors'][0]['message'])

        results = search_results['data']['Get'][self.classNm]
        return results

def search_do(input_, alp):
    searcher = WeaviateSemanticSearch(vdb)
    results = searcher.hybrid_search(input_, 3, alpha=alp)

    result_li = []
    for _, result in enumerate(results, 1):
        result_li.append(result['content'])

    return result_li

if __name__ == "__main__":
    client = WeaviateSemanticSearch(vdb)

    # 統計筆數
    count_result = client.aggregate_count()
    print(count_result)

    # 輸出所有資料
    data_result = client.get_all_data()
    print(data_result)

    # 刪除此向量庫
    # client.delete_class()

    # An alpha of 1 is a pure vector search
    # An alpha of 0 is a pure keyword search
    quest = "老師 蔡炎龍 學生"
    with open("result.txt", 'w', encoding='utf-8') as file:
        file.write(str(search_do(quest, alp=1)) + "\n\n\n" + str(search_do(quest, alp=0.1)))
