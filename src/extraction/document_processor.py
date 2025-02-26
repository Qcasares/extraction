"""Document processor module."""

import json
import logging
import os
import xml.etree.ElementTree as ET
from csv import DictWriter
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

import docx
import requests
import spacy
from docx.document import Document as DocxDocument
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class DocumentMetadata:
    """Metadata extracted from a document."""

    title: str = ""
    author: str = ""
    created_date: str = ""
    last_modified_date: str = ""
    file_path: str = ""
    filename: str = ""


@dataclass
class DocumentSection:
    """A section of a document with heading and content."""

    title: str
    level: int
    content: str
    entities: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class ProcessedDocument:
    """Document with extracted and processed content."""

    metadata: DocumentMetadata
    sections: List[DocumentSection] = field(default_factory=list)
    raw_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            "metadata": asdict(self.metadata),
            "sections": [asdict(section) for section in self.sections],
            "raw_text": self.raw_text,
        }


class DocumentProcessor:
    """Processor for Word documents."""

    def __init__(
        self,
        input_directory: str,
        output_format: str = "json",
        use_ai: bool = False,
        ollama_model: Optional[str] = None,
        ollama_api_base: str = "http://localhost:11434/api",
    ) -> None:
        """Initialize the document processor.

        Args:
            input_directory: Directory to process documents from
            output_format: Output format (json, xml, csv)
            use_ai: Whether to use AI for analysis
            ollama_model: Ollama model to use
            ollama_api_base: Base URL for Ollama API
        """
        self.input_directory = Path(input_directory)
        self.output_format = output_format.lower()
        self.use_ai = use_ai
        self.ollama_model = ollama_model
        self.ollama_api_base = ollama_api_base
        self.nlp = spacy.load("en_core_web_sm")
        self.processed_documents: List[ProcessedDocument] = []

        # Validate input directory
        if not self.input_directory.exists():
            raise ValueError(f"Input directory does not exist: {input_directory}")
        if not self.input_directory.is_dir():
            raise ValueError(f"Input path is not a directory: {input_directory}")

        # Validate output format
        if self.output_format not in ("json", "xml", "csv"):
            raise ValueError(
                f"Invalid output format: {output_format}. Must be json, xml, or csv."
            )

        logger.info(
            "Initialized DocumentProcessor with input_directory=%s, output_format=%s, use_ai=%s",
            input_directory,
            output_format,
            use_ai,
        )

    def find_documents(self) -> List[Path]:
        """Find all Word documents in the input directory.

        Returns:
            List of paths to Word documents
        """
        logger.info("Searching for .docx files in %s", self.input_directory)
        docx_files: List[Path] = []

        for path in self.input_directory.glob("**/*.docx"):
            if path.is_file():
                docx_files.append(path)

        logger.info("Found %d .docx files", len(docx_files))
        return docx_files

    def extract_metadata(self, doc: DocxDocument, file_path: Path) -> DocumentMetadata:
        """Extract metadata from a Word document.

        Args:
            doc: Word document
            file_path: Path to the document

        Returns:
            Document metadata
        """
        properties = doc.core_properties
        metadata = DocumentMetadata(
            title=properties.title if properties.title else "",
            author=properties.author if properties.author else "",
            created_date=properties.created.isoformat() if properties.created else "",
            last_modified_date=(
                properties.modified.isoformat() if properties.modified else ""
            ),
            file_path=str(file_path.absolute()),
            filename=file_path.name,
        )
        return metadata

    def extract_text(self, doc: DocxDocument) -> str:
        """Extract text from a Word document.

        Args:
            doc: Word document

        Returns:
            Extracted text
        """
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)

        # Extract text from tables
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    row_text.append(cell.text)
                full_text.append(" | ".join(row_text))

        return "\n".join(full_text)

    def identify_sections(self, doc: DocxDocument) -> List[DocumentSection]:
        """Identify document sections based on headings.

        Args:
            doc: Word document

        Returns:
            List of document sections
        """
        sections: List[DocumentSection] = []
        current_section: Optional[DocumentSection] = None
        content_buffer: List[str] = []

        for paragraph in doc.paragraphs:
            # Check if paragraph is a heading
            # Check if paragraph has a style and if it's a heading
            if paragraph.style and hasattr(paragraph.style, 'name') and paragraph.style.name and paragraph.style.name.startswith("Heading"):
                # Extract heading level (Heading 1, Heading 2, etc.)
                try:
                    level = int(paragraph.style.name.split()[-1])
                except ValueError:
                    level = 0

                # If we have a previous section, save its content
                if current_section is not None:
                    current_section.content = "\n".join(content_buffer)
                    sections.append(current_section)

                # Create a new section
                current_section = DocumentSection(
                    title=paragraph.text,
                    level=level,
                    content="",
                )
                content_buffer = []
            elif current_section is not None:
                # Add paragraph to current section
                if paragraph.text.strip():
                    content_buffer.append(paragraph.text)
            else:
                # Text before any heading - create a default section
                if paragraph.text.strip() and not current_section:
                    current_section = DocumentSection(
                        title="Introduction",
                        level=0,
                        content="",
                    )
                    content_buffer.append(paragraph.text)

        # Don't forget the last section
        if current_section is not None:
            current_section.content = "\n".join(content_buffer)
            sections.append(current_section)

        return sections

    def analyze_content(self, processed_doc: ProcessedDocument) -> None:
        """Analyze document content to extract entities and structure.

        Args:
            processed_doc: Processed document to analyze
        """
        # Process each section with spaCy
        for section in processed_doc.sections:
            if not section.content:
                continue

            doc = self.nlp(section.content)

            # Extract named entities
            entities = []
            for ent in doc.ents:
                entities.append(
                    {
                        "text": ent.text,
                        "label": ent.label_,
                        "start": ent.start_char,
                        "end": ent.end_char,
                    }
                )

            section.entities = entities

    def analyze_with_ai(self, processed_doc: ProcessedDocument) -> None:
        """Use Ollama AI to enhance document analysis.

        Args:
            processed_doc: Processed document to analyze
        """
        if not self.use_ai or not self.ollama_model:
            return

        try:
            # Prepare the prompt with document content
            prompt = f"""
            Please analyze this document and identify key sections, entities, and structure.
            
            Document Title: {processed_doc.metadata.title}
            Document Content:
            {processed_doc.raw_text[:1000]}...
            
            Provide a structured analysis with:
            1. Key topics
            2. Main entities mentioned
            3. Relationships between entities
            4. Document structure
            """

            # Call Ollama API
            response = requests.post(
                f"{self.ollama_api_base}/generate",
                json={
                    "model": self.ollama_model,
                    "prompt": prompt,
                    "stream": False,
                },
                timeout=30,
            )

            if response.status_code == 200:
                result = response.json()
                # Process AI response here - actual implementation would depend on how
                # you want to use the AI analysis
                logger.info(
                    "AI analysis completed for %s", processed_doc.metadata.filename
                )
            else:
                logger.warning(
                    "Failed to get AI analysis: %s, %s",
                    response.status_code,
                    response.text,
                )
        except Exception as e:
            logger.error("Error during AI analysis: %s", str(e))
            # Fall back to rule-based analysis
            logger.info("Falling back to rule-based analysis")

    def process_document(self, file_path: Path) -> ProcessedDocument:
        """Process a single document.

        Args:
            file_path: Path to the document

        Returns:
            Processed document
        """
        logger.info("Processing document: %s", file_path)

        try:
            doc = docx.Document(str(file_path))

            # Extract metadata
            metadata = self.extract_metadata(doc, file_path)

            # Extract raw text
            raw_text = self.extract_text(doc)

            # Identify sections
            sections = self.identify_sections(doc)

            # Create processed document
            processed_doc = ProcessedDocument(
                metadata=metadata,
                sections=sections,
                raw_text=raw_text,
            )

            # Analyze content
            self.analyze_content(processed_doc)

            # Use AI for enhanced analysis if enabled
            if self.use_ai and self.ollama_model:
                self.analyze_with_ai(processed_doc)

            return processed_doc

        except Exception as e:
            logger.error("Error processing document %s: %s", file_path, str(e))
            # Return an empty document with error information
            return ProcessedDocument(
                metadata=DocumentMetadata(
                    file_path=str(file_path),
                    filename=file_path.name,
                )
            )

    def process_all(self) -> List[ProcessedDocument]:
        """Process all documents in the input directory.

        Returns:
            List of processed documents
        """
        documents = self.find_documents()
        self.processed_documents = []

        for file_path in tqdm(documents, desc="Processing documents"):
            processed_doc = self.process_document(file_path)
            self.processed_documents.append(processed_doc)

        return self.processed_documents

    def save_results(self, output_path: str) -> None:
        """Save processed results to the specified path.

        Args:
            output_path: Path to save results to
        """
        output_dir = Path(output_path)
        output_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if self.output_format == "json":
            self._save_as_json(output_dir, timestamp)
        elif self.output_format == "xml":
            self._save_as_xml(output_dir, timestamp)
        elif self.output_format == "csv":
            self._save_as_csv(output_dir, timestamp)

        logger.info(
            "Saved %d documents to %s in %s format",
            len(self.processed_documents),
            output_dir,
            self.output_format,
        )

    def _save_as_json(self, output_dir: Path, timestamp: str) -> None:
        """Save results as JSON.

        Args:
            output_dir: Directory to save to
            timestamp: Timestamp for filename
        """
        # Save individual documents
        for doc in self.processed_documents:
            filename = f"{doc.metadata.filename.replace('.docx', '')}_{timestamp}.json"
            with open(output_dir / filename, "w") as f:
                json.dump(doc.to_dict(), f, indent=2)

        # Save summary file with all documents
        summary_file = output_dir / f"document_summary_{timestamp}.json"
        with open(summary_file, "w") as f:
            json.dump([doc.to_dict() for doc in self.processed_documents], f, indent=2)

    def _save_as_xml(self, output_dir: Path, timestamp: str) -> None:
        """Save results as XML.

        Args:
            output_dir: Directory to save to
            timestamp: Timestamp for filename
        """
        for doc in self.processed_documents:
            filename = f"{doc.metadata.filename.replace('.docx', '')}_{timestamp}.xml"

            # Create XML structure
            root = ET.Element("document")

            # Add metadata
            metadata_elem = ET.SubElement(root, "metadata")
            for key, value in asdict(doc.metadata).items():
                ET.SubElement(metadata_elem, key).text = str(value)

            # Add sections
            sections_elem = ET.SubElement(root, "sections")
            for section in doc.sections:
                section_elem = ET.SubElement(sections_elem, "section")
                ET.SubElement(section_elem, "title").text = section.title
                ET.SubElement(section_elem, "level").text = str(section.level)
                ET.SubElement(section_elem, "content").text = section.content

                # Add entities
                entities_elem = ET.SubElement(section_elem, "entities")
                for entity in section.entities:
                    entity_elem = ET.SubElement(entities_elem, "entity")
                    for key, value in entity.items():
                        ET.SubElement(entity_elem, key).text = str(value)

            # Write to file
            tree = ET.ElementTree(root)
            # Only use indent if available (Python 3.9+)
            if hasattr(ET, "indent"):
                ET.indent(tree, space="  ")
            tree.write(output_dir / filename, encoding="utf-8", xml_declaration=True)

    def _save_as_csv(self, output_dir: Path, timestamp: str) -> None:
        """Save results as CSV.

        Args:
            output_dir: Directory to save to
            timestamp: Timestamp for filename
        """
        # CSV for document metadata
        metadata_file = output_dir / f"document_metadata_{timestamp}.csv"
        with open(metadata_file, "w", newline="") as f:
            fieldnames = [
                "filename",
                "title",
                "author",
                "created_date",
                "last_modified_date",
                "file_path",
            ]
            writer = DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for doc in self.processed_documents:
                writer.writerow(asdict(doc.metadata))

        # CSV for sections
        sections_file = output_dir / f"document_sections_{timestamp}.csv"
        with open(sections_file, "w", newline="") as f:
            fieldnames = ["document_filename", "section_title", "level", "content"]
            writer = DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for doc in self.processed_documents:
                for section in doc.sections:
                    writer.writerow(
                        {
                            "document_filename": doc.metadata.filename,
                            "section_title": section.title,
                            "level": section.level,
                            "content": section.content,
                        }
                    )

        # CSV for entities
        entities_file = output_dir / f"document_entities_{timestamp}.csv"
        with open(entities_file, "w", newline="") as f:
            fieldnames = [
                "document_filename",
                "section_title",
                "entity_text",
                "entity_label",
                "entity_start",
                "entity_end",
            ]
            writer = DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for doc in self.processed_documents:
                for section in doc.sections:
                    for entity in section.entities:
                        writer.writerow(
                            {
                                "document_filename": doc.metadata.filename,
                                "section_title": section.title,
                                "entity_text": entity["text"],
                                "entity_label": entity["label"],
                                "entity_start": entity["start"],
                                "entity_end": entity["end"],
                            }
                        )
