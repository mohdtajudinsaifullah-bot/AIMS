import os
import time
from dotenv import load_dotenv
from pinecone import Pinecone, ServerlessSpec
from langchain_openai import OpenAIEmbeddings
from langchain_pinecone import PineconeVectorStore
from langchain_core.documents import Document

# 1. Muat turun API Keys dari fail .env
load_dotenv()
pinecone_key = os.getenv("PINECONE_API_KEY")
openai_key = os.getenv("OPENAI_API_KEY")

# 2. Tetapkan nama index untuk projek ni (huruf kecil & sengkang sahaja)
nama_index = "projek-aims"

# 3. Hubungkan ke Pinecone
pc = Pinecone(api_key=pinecone_key)

print(f"Menyemak status index '{nama_index}' di Pinecone...")

# 4. Semak dan cipta index jika belum wujud
if nama_index not in pc.list_indexes().names():
    print(f"Mencipta index baru: '{nama_index}' (proses ni ambil masa sikit)...")
    pc.create_index(
        name=nama_index,
        dimension=1536, # Dimensi vektor untuk model OpenAI text-embedding-3-small
        metric="cosine",
        spec=ServerlessSpec(
            cloud="aws",
            region="us-east-1" # Boleh tukar ikut region percuma Pinecone kau
        )
    )
    # Tunggu sehingga index sedia digunakan
    while not pc.describe_index(nama_index).status['ready']:
        time.sleep(1)
    print("Index baru berjaya dicipta!")
else:
    print(f"Index '{nama_index}' dah sedia wujud. Kita guna yang tu je.")

# 5. Tetapkan model Embedding dari OpenAI
print("Memulakan OpenAI Embeddings...")
embeddings = OpenAIEmbeddings(model="text-embedding-3-small")

# 6. Sediakan data contoh (Boleh ganti dengan data PDF/Teks kau nanti)
print("Menyediakan data untuk diserap...")
data_contoh = [
    Document(
        page_content="Projek AIMS adalah satu inisiatif sistem pengurusan dokumen pintar.", 
        metadata={"kategori": "pengenalan", "sumber": "dokumen_utama"}
    ),
    Document(
        page_content="Pinecone bertindak sebagai pangkalan data vektor untuk carian pantas.", 
        metadata={"kategori": "teknologi", "sumber": "wiki_sistem"}
    ),
    Document(
        page_content="Langchain digunakan sebagai jambatan antara OpenAI dan Pinecone.", 
        metadata={"kategori": "teknologi", "sumber": "wiki_sistem"}
    )
]

# 7. Masukkan data ke dalam Pinecone (Proses Serapan/Indexing)
print("Sedang menyerap (indexing) data ke dalam Pinecone...")
vectorstore = PineconeVectorStore.from_documents(
    data_contoh,
    index_name=nama_index,
    embedding=embeddings
)

print("\nAlhamdulillah! Data berjaya diserap ke dalam Pinecone. Projek AIMS on track!")