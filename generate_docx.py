import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

def create_cv(version_type, output_path):
    doc = docx.Document()
    
    # Page setup - Margins 0.5 in for clean fit
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        section.page_width = Inches(8.27)  # A4
        section.page_height = Inches(11.69)
        
    # Default Normal style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(2)
    normal_style.paragraph_format.space_before = Pt(0)
    
    def add_p(text='', bold=False, italic=False, font_size=9.5, color_rgb=(30, 41, 59), align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=2):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(*color_rgb)
        return p

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(15, 23, 42)
        # Bottom border for professional ATS divider
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="8" w:space="2" w:color="0F172A"/></w:pBdr>')
        pPr.append(pBdr)
        return p

    def add_item_header(left_title, right_meta, subtitle='', tech=''):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        
        r1 = p.add_run(left_title)
        r1.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(15, 23, 42)
        
        if right_meta:
            r_meta = p.add_run(f'  |  {right_meta}')
            r_meta.bold = False
            r_meta.font.size = Pt(9.5)
            r_meta.font.color.rgb = RGBColor(71, 85, 105)
            
        if subtitle or tech:
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(2)
            p2.paragraph_format.keep_with_next = True
            if subtitle:
                r_sub = p2.add_run(subtitle)
                r_sub.bold = True
                r_sub.font.size = Pt(9.5)
                r_sub.font.color.rgb = RGBColor(2, 132, 199)
            if subtitle and tech:
                r_sep = p2.add_run(' — ')
                r_sep.font.size = Pt(9.5)
                r_sep.font.color.rgb = RGBColor(100, 116, 139)
            if tech:
                r_tech = p2.add_run(f'Tech: {tech}')
                r_tech.italic = True
                r_tech.font.size = Pt(9)
                r_tech.font.color.rgb = RGBColor(71, 85, 105)

    def add_bullet(lead_bold, rest_text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Inches(0.25)
        if lead_bold:
            r_bold = p.add_run(lead_bold)
            r_bold.bold = True
            r_bold.font.size = Pt(9.5)
            r_bold.font.color.rgb = RGBColor(15, 23, 42)
        r_rest = p.add_run(rest_text)
        r_rest.font.size = Pt(9.5)
        r_rest.font.color.rgb = RGBColor(30, 41, 59)
        return p

    # --- HEADER ---
    add_p('MOHAMED MOSTAFA HELAL', bold=True, font_size=18, color_rgb=(15, 23, 42), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=1)
    
    headline_text = 'AI & Machine Learning Engineer | LLMs, Agentic Systems & Computer Vision'
    add_p(headline_text, bold=True, font_size=10.5, color_rgb=(2, 132, 199), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=3)
    
    if version_type == 'Egypt':
        contact_line_1 = 'Giza, Egypt  |  +20 103 046 9659  |  mohamed200031921@gmail.com  |  Military Status: Exempted'
    else:
        contact_line_1 = 'Egypt (Remote, Flexible Timezone)  |  +20 103 046 9659  |  mohamed200031921@gmail.com'
    contact_line_2 = 'LinkedIn: linkedin.com/in/mohamed-mostafa-454b70271  |  GitHub: github.com/DarkStrox  |  Hugging Face: huggingface.co/AitroniXAI'
    
    add_p(contact_line_1, font_size=9, color_rgb=(71, 85, 105), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=1)
    add_p(contact_line_2, font_size=9, color_rgb=(71, 85, 105), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4)

    # --- SUMMARY ---
    add_section_header('Professional Summary')
    summary_text = (
        'AI & Machine Learning Engineer and 2026 Computer Science graduate (GPA 3.62/4.0, Honors) specialized in LLM agentic systems, edge inference, and real-time computer vision. '
        'Architected AitroniX 1s, an open-source 2.6B bilingual reasoning SLM achieving 68.75% strict AST match on BFCL v4, optimized for mobile edge deployment via ~1.56 GB GGUF. '
        'Proven hands-on track record building local hybrid RAG architectures (FAISS + BM25) and deploying hardware-accelerated multi-stream video analytics pipelines running at 30+ FPS.'
    )
    add_p(summary_text, font_size=9.5, space_before=2, space_after=4)

    # --- EXPERIENCE ---
    add_section_header('Professional Experience')
    
    # Aitronix
    add_item_header('Aitronix', 'Jan 2026 – Present | Giza, Egypt' if version_type == 'Egypt' else 'Jan 2026 – Present | Remote', subtitle='AI Engineer')
    add_bullet('AitroniX 1s Model Fine-Tuning: ', 'Fine-tuned and released AitroniX 1s, an open-source 2.6B bilingual (Arabic/English) reasoning SLM utilizing a hybrid Conv+Attention architecture with a 128K context window using Unsloth and LoRA/QLoRA.')
    add_bullet('Structured Tool-Calling Evaluation: ', 'Evaluated structured tool calling on Berkeley Function-Calling Leaderboard (BFCL v4, 400 problems), achieving a 68.75% strict AST match and 100% valid tool syntax (vs. Qwen3.5-9B at 60.13% and Gemma-4-8B at 48.39%).')
    add_bullet('Edge & High-Throughput Deployment: ', 'Quantized and optimized model weights into Q4_K_M GGUF (~1.56 GB memory footprint) for private on-device mobile execution and configured high-throughput server deployment using vLLM and SGLang.')
    add_bullet('Real-Time Video Analytics: ', 'Engineered hardware-accelerated computer vision pipelines and multi-camera RTSP streaming engines, sustaining 30+ FPS across concurrent live feeds via OpenCL GPU acceleration and asynchronous worker queues.')

    # Drug Control Fund
    add_item_header('Drug Control & Addiction Prevention Fund', 'Jan 2024 – Dec 2025 | Cairo, Egypt' if version_type == 'Egypt' else 'Jan 2024 – Dec 2025 | Egypt', subtitle="Full-Stack Developer — Mutatawi' Rasmi")
    add_bullet('End-to-End Platform Development: ', "Developed the Mutatawi' Rasmi web application ecosystem from scratch, implementing responsive user interfaces, backend business logic, and relational database schemas to coordinate nationwide anti-drug youth initiatives.")
    add_bullet('Geospatial Hotspot Analytics: ', 'Engineered data-driven geospatial hotspot analytics and risk visualization maps across Egyptian governorates, empowering administrators to track field missions and allocate operational resources.')
    add_bullet('Volunteer Operations & Engagement: ', 'Implemented volunteer engagement engines, automated leaderboards, milestone verification workflows, and mission tracking modules to streamline volunteer coordination and reporting.')

    # Samsung Innovation Campus
    add_item_header('Samsung Innovation Campus', 'Jul 2024 – Oct 2024 | Giza, Egypt' if version_type == 'Egypt' else 'Jul 2024 – Oct 2024 | Egypt', subtitle='AI & Data Science Resident Intern')
    add_bullet('Quantitative Feature Engineering: ', 'Co-developed a quantitative time-series forecasting framework by processing raw market data and engineering 15 financial features, including rolling standard deviations, momentum indices, and historical lags.')
    add_bullet('Predictive Modeling & Optimization: ', 'Constructed and tuned ensemble regression and time-series models (XGBoost, Random Forest, and SARIMA), systematically reducing Mean Absolute Percentage Error (MAPE) from 8.3% to 6.5%.')
    add_bullet('Exploratory Data Analysis & Dashboards: ', 'Conducted extensive exploratory data analysis (EDA) on customer retention datasets using Pandas and designed interactive dashboards with Plotly and Seaborn to track operational attrition trends.')

    # --- PROJECTS ---
    add_section_header('Technical Projects')

    # Project 1: AitroniX 1s
    add_item_header('AitroniX 1s — Bilingual 2.6B Agentic Reasoning SLM', 'Jan 2026', tech='Hybrid Conv+Attention, Unsloth, LoRA/QLoRA, GGUF, vLLM, SGLang, BFCL v4')
    add_bullet('Architecture & Open-Source Release: ', 'Architected and trained an open-source 2.6B bilingual (Arabic/English) reasoning model featuring a hybrid Convolution + Attention architecture, 128K context window, and native <think> Chain-of-Thought reasoning.')
    add_bullet('Benchmark Performance: ', 'Evaluated on Berkeley Function-Calling Leaderboard (BFCL v4 AST, all 400 problems), achieving 68.75% strict AST match with 100% valid tool syntax (compared to Qwen3.5-9B at 60.13%, Gemma-4-8B at 48.39%, and Gemma-4-2B at 35.98%).')
    add_bullet('Evaluation & Quantization: ', 'Benchmarked 77.83 on ToolSandbox, 85.49% on IFStruct, 51.87% on AIME25 math, and quantized weights to Q4_K_M GGUF (~1.56 GB) for zero-latency private inference on mobile edge devices.')

    # Project 2: HyperRAG
    add_item_header('HyperRAG — Local Hybrid RAG Engine', '2026', tech='FAISS, BM25, Reciprocal Rank Fusion (RRF), llama.cpp, Gemma 4 E4B, Python')
    add_bullet('Hybrid Retrieval Pipeline: ', 'Engineered a local hybrid retrieval pipeline combining dense vector search (FAISS) and sparse lexical keyword search (BM25) via Reciprocal Rank Fusion (RRF), eliminating keyword entity retrieval misses.')
    add_bullet('Private Offline Inference: ', 'Deployed local LLM generation via llama.cpp using Gemma 4 E4B for private, zero-cloud inference, featuring dynamic 800-token chunking (120-token overlap) and verifiable document citation tracking.')

    # Project 3: Face Recognition Engine
    add_item_header('Aitronix Real-Time High-FPS Face Recognition Engine', '2026', tech='Python, OpenCV, YuNet, SFace, OpenCL, Multithreading, RTSP, LibJPEG')
    add_bullet('High-Speed Face Recognition: ', 'Built a real-time face recognition pipeline utilizing YuNet for single-pass tensor face detection (~2.5ms inference via OpenCL GPU acceleration) and SFace for 128D embeddings with Cosine Similarity vector matching.')
    add_bullet('Multithreading & Byte Decoding: ', 'Designed non-blocking asynchronous multithreading with decoupled worker queues delivering sustained 30+ FPS; implemented custom native LibJPEG byte decoder to eliminate network buffering delay over RTSP and USB feeds.')

    # Project 4: CamScence
    add_item_header('CamScence — Multi-Camera Edge Traffic Analytics System', '2026', tech='YOLO (YOLO26n), RTSP Multi-Streaming, Speed Estimation, Spatial Heatmaps, Python')
    add_bullet('Multi-Camera Edge Pipeline: ', 'Architected an edge video analytics system concurrently ingesting 11+ RTSP IP camera streams with dynamic multi-grid switching and automated feed monitoring.')
    add_bullet('Tracking & Speed Estimation: ', 'Implemented real-time vehicle detection, persistent multi-object tracking, automated vehicle speed estimation (km/h), and spatial heatmap overlays for traffic congestion and bottleneck analysis.')

    # Project 5: Surveillance Threat Detection
    add_item_header('Real-Time Video Surveillance Threat & Burglary Detection', '2025 – 2026', tech='Microsoft Florence-2, OpenAI CLIP, OpenCV, Zero-Shot Grounding, Python')
    add_bullet('Multimodal Video Analysis: ', 'Developed an automated video surveillance pipeline combining Vision-Language Models (Microsoft Florence-2, OpenAI CLIP) and OpenCV to detect security incidents, burglary, and physical violence.')
    add_bullet('Zero-Shot Alerting: ', 'Applied zero-shot visual classification and phrase grounding to score threat probabilities frame-by-frame and trigger dynamic real-time on-screen security alerts.')

    # Project 6: YOLO-World
    add_item_header('Open-Vocabulary Real-Time Detection & Tracking (YOLO-World)', '2025 – 2026', tech='YOLO-World, Open-Vocabulary Vision, PyTorch, Real-Time GPU Tracking')
    add_bullet('Promptable Zero-Shot Tracking: ', 'Deployed YOLO-World to detect and track custom arbitrary natural language prompted visual concepts without model retraining or additional labeled data.')
    add_bullet('High-FPS GPU Execution: ', 'Optimized local GPU inference pipelines to execute real-time object tracking at 70+ FPS on high-speed video feeds while maintaining continuous spatial ID tracking.')

    # Project 7: Arabian Horses Graduation Project
    add_item_header('Integrated AI Ecosystem for Arabian Horses (Graduation Project)', 'Sep 2025 – Jun 2026', tech='React.js, ASP.NET Core Web API, SQL Server, Python, ConvNeXt, PyTorch')
    add_bullet('Full-Stack Platform: ', 'Architected an end-to-end multi-role web platform (React frontend, ASP.NET Core REST API backend, SQL Server) for horse auctions, farm management (Marabet), and medical record tracking.')
    add_bullet('Deep Learning Pedigree & Valuation: ', 'Integrated a Python deep learning pipeline utilizing ConvNeXt for horse pedigree and strain identification alongside regression models for market valuation and price estimation (Grade: Excellent).')

    # Other Projects
    p_other = doc.add_paragraph()
    p_other.paragraph_format.space_before = Pt(3)
    p_other.paragraph_format.space_after = Pt(3)
    r_other_lbl = p_other.add_run('Additional Projects: ')
    r_other_lbl.bold = True
    r_other_lbl.font.size = Pt(9)
    r_other_lbl.font.color.rgb = RGBColor(15, 23, 42)
    r_other_txt = p_other.add_run('NLP Benchmark (DistilBERT 91.36% on 50k IMDb reviews vs RNN/LSTM baselines with FP16); WhatsApp AI Assistant (n8n + Supabase pgvector + WhatsApp Cloud API); Moondream Edge VLM (webcam VQA); Plant Disease Classifier (38-class CNN, 95.4% accuracy).')
    r_other_txt.font.size = Pt(9)
    r_other_txt.font.color.rgb = RGBColor(71, 85, 105)

    # --- SKILLS ---
    add_section_header('Technical Skills')
    def add_skill_line(cat, skills):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r_cat = p.add_run(f'{cat}: ')
        r_cat.bold = True
        r_cat.font.size = Pt(9.5)
        r_cat.font.color.rgb = RGBColor(15, 23, 42)
        r_sk = p.add_run(skills)
        r_sk.font.size = Pt(9.5)
        r_sk.font.color.rgb = RGBColor(51, 65, 85)

    add_skill_line('LLMs & Agentic Systems', 'LLM Fine-Tuning (Unsloth, LoRA/QLoRA), Structured Tool Calling (JSON/XML, AST), Function Calling, GGUF Quantization, vLLM, SGLang, llama.cpp, Ollama, Hugging Face Transformers, Tokenization, Prompt Engineering')
    add_skill_line('Retrieval & Search (RAG)', 'Hybrid RAG, Dense Vector Search (FAISS), Sparse Lexical Indexing (BM25), Reciprocal Rank Fusion (RRF), Vector Databases (pgvector / Supabase)')
    add_skill_line('Computer Vision & Multimodal', 'YOLO (v8, YOLO-World, YOLO26n), Vision-Language Models (Florence-2, CLIP, Moondream), Face Recognition (YuNet, SFace), ConvNeXt, 3D CNNs, MobileNet, Transfer Learning, OpenCV, OpenCL GPU Acceleration, RTSP Multi-Streaming, Spatial Heatmap Tracking & Speed Estimation')
    add_skill_line('Machine Learning & Data Science', 'PyTorch, TensorFlow, Keras, Scikit-learn, XGBoost, Random Forest, SARIMA, Time-Series Forecasting, Feature Engineering, Pandas, NumPy, Plotly, Seaborn')
    add_skill_line('Backend & Engineering', 'Python, C#, ASP.NET Core Web API, React.js, JavaScript, PostgreSQL, Microsoft SQL Server, RESTful APIs, Asynchronous Multithreading, n8n Automation, Docker, Git, GitHub')
    add_skill_line('Languages', 'Arabic (Native), English (Professional / Fluent)')

    # --- EDUCATION ---
    add_section_header('Education')
    add_item_header('Misr University for Science and Technology (MUST)', 'Giza, Egypt | Sep 2022 – Jun 2026', subtitle='Bachelor of Science in Computer Science, Major in Artificial Intelligence')
    add_bullet('Academic Distinction: ', 'Cumulative GPA: 3.62 / 4.0 (Excellent with Honors)')
    add_bullet('Graduation Project: ', 'Integrated AI Ecosystem for Arabian Horses (Grade: Excellent)')

    # --- CERTIFICATIONS ---
    add_section_header('Certifications & Professional Credentials')
    add_bullet('Samsung Innovation Campus Certified Professional in AI & Data Science (2024): ', 'Issued by Samsung Electronics & Ministry of Communications and Information Technology (MCIT).')
    add_bullet('Deep Learning & Computer Vision Specialization: ', 'Comprehensive study in CNN architectures, object detection, and transfer learning.')
    add_bullet('Machine Learning & Advanced Statistical Modeling: ', 'Time-series forecasting, regression, and ensemble models.')

    doc.save(output_path)
    print(f'Successfully generated: {output_path}')

if __name__ == '__main__':
    create_cv('Egypt', 'e:/Aitronix Projects/cv_preview/Mohamed_Mostafa_Helal_CV_Egypt.docx')
    create_cv('International', 'e:/Aitronix Projects/cv_preview/Mohamed_Mostafa_Helal_CV_International.docx')
