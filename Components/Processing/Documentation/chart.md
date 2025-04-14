flowchart TD
    PDF[PDF Document] --> PDFProcessor[pdf_processor.py]
    
    subgraph Processing Package
        Document[document.py]
        subgraph Core
            Extraction[extraction.py]
        end
        subgraph Parsers
            TableParser[table.py]
            KeywordParser[keywords.py]
        end
        subgraph Utilities
            TextUtil[text.py]
            Cleaner[cleaner.py]
            Merger[merger.py]
        end
    end
    
    subgraph Components
        GenInfo[GeneralInfo.py]
        PDFExtract[pdf_extractor.py]
        BusRules[business_rules.py]
    end
    
    PDFProcessor --> Document
    Document --> Extraction
    Document --> Merger
    
    Extraction --> GenInfo
    GenInfo --> PDFExtract
    PDFExtract --> BusRules
    
    Extraction --> TableParser
    Extraction --> KeywordParser
    Extraction --> TextUtil
    Extraction --> Cleaner
    
    Document --> JSON[JSON Output]
    
    %% Styling
    classDef main fill:#f96,stroke:#333,stroke-width:2px;
    classDef processing fill:#9cf,stroke:#333,stroke-width:1px;
    classDef components fill:#fcf,stroke:#333,stroke-width:1px;
    classDef parsers fill:#cfc,stroke:#333,stroke-width:1px;
    classDef utilities fill:#ff9,stroke:#333,stroke-width:1px;
    classDef io fill:#eee,stroke:#333,stroke-width:2px;
    
    class PDF,JSON io;
    class PDFProcessor main;
    class Document,Extraction processing;
    class GenInfo,PDFExtract,BusRules components;
    class TableParser,KeywordParser parsers;
    class TextUtil,Cleaner,Merger utilities;