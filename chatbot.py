import os
from langchain_community.document_loaders import TextLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.chains import create_retrieval_chain
from langchain.chains.combine_documents import create_stuff_documents_chain
from langchain_core.prompts import ChatPromptTemplate

# 1. Setup the AI (You will need to set an environment variable named GOOGLE_API_KEY)
llm = ChatGoogleGenerativeAI(model="gemini-1.5-flash", temperature=0.2)

# 2. Setup the free local embedding model
embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")

def initialize_knowledge_base():
    # 3. Load the manual
    loader = TextLoader("manual.txt")
    docs = loader.load()

    # 4. Chop the manual into small, searchable chunks
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    splits = text_splitter.split_documents(docs)

    # 5. Build the searchable database
    vectorstore = FAISS.from_documents(splits, embeddings)
    return vectorstore.as_retriever()

# Initialize it once when the server starts
retriever = initialize_knowledge_base()

def ask_timegen_bot(user_question):
    # 6. Give the AI its strict instructions
    system_prompt = (
        "You are the TimeGen Assistant, an expert on the academic scheduling software."
        "Use the provided context to answer the user's question."
        "If the answer is not in the context, say 'I can only answer questions about the TimeGen software based on the manual.'"
        "Keep your answers concise, friendly, and formatted with line breaks if necessary."
        "\n\n"
        "{context}"
    )

    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", "{input}"),
    ])

    question_answer_chain = create_stuff_documents_chain(llm, prompt)
    rag_chain = create_retrieval_chain(retriever, question_answer_chain)

    response = rag_chain.invoke({"input": user_question})
    return response["answer"]