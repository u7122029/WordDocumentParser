# API Reference

Full API documentation for the WordDocumentParser library, auto-generated from XML documentation comments.

## Core Classes

| Class | Description |
|-------|-------------|
| @WordDocumentParser.WordDocument | Primary document wrapper with property access and content tree |
| @WordDocumentParser.DocumentNode | Tree node with type, text, formatting, children, and parent reference |
| @WordDocumentParser.WordDocumentTreeParser | Parses .docx files into the tree model |
| @WordDocumentParser.WordDocumentTreeWriter | Writes the tree model back to .docx with formatting preservation |

## Extension Methods

| Extension Class | Purpose |
|----------------|---------|
| @WordDocumentParser.Extensions.TableExtensions | Cell access, structural modification, formatting, nested tables |
| @WordDocumentParser.Extensions.DocumentMergeExtensions | Append, concatenate, extract sections, insert nodes, clone |
| @WordDocumentParser.Extensions.ContentControlExtensions | Find, update, remove SDT controls |
| @WordDocumentParser.Extensions.FontExtensions | Paragraph/document/range font changes, font queries |
| @WordDocumentParser.Extensions.StyleExtensions | Find by style, change styles, style distribution |
| @WordDocumentParser.Extensions.TreeNavigationExtensions | FindAll, GetPath, GetHeadingPath, siblings, flatten |
| @WordDocumentParser.Extensions.TreeQueryExtensions | GetAllHeadings, GetAllTables, GetTableOfContents, CountByType |
| @WordDocumentParser.Extensions.DocumentPropertyExtensions | Property field queries, metadata text extraction |
| @WordDocumentParser.Extensions.SerializationExtensions | SaveToFile, SaveToStream, ToDocxBytes |

## Models

### Formatting

| Class | Description |
|-------|-------------|
| @WordDocumentParser.Models.Formatting.FormattedRun | A text run with associated formatting |
| @WordDocumentParser.Models.Formatting.RunFormatting | Font, bold, italic, underline, color, and other run-level properties |
| @WordDocumentParser.Models.Formatting.ParagraphFormatting | Alignment, spacing, indentation, and other paragraph-level properties |
| @WordDocumentParser.Models.Formatting.TableFormatting | Table-level formatting (alignment, width, borders) |
| @WordDocumentParser.Models.Formatting.TableCellFormatting | Cell-level formatting (shading, vertical alignment, borders) |
| @WordDocumentParser.Models.Formatting.TableRowFormatting | Row-level formatting (height, header repeat) |
| @WordDocumentParser.Models.Formatting.BorderFormatting | Border style, size, color, and spacing |

### Content Controls

| Class | Description |
|-------|-------------|
| @WordDocumentParser.Models.ContentControls.ContentControlProperties | Full metadata for an SDT control |
| @WordDocumentParser.Models.ContentControls.ContentControlType | Enum: PlainText, RichText, DropDownList, ComboBox, Date, Checkbox, Picture |
| @WordDocumentParser.Models.ContentControls.ContentControlListItem | Display text and value for dropdown/combobox items |

### Tables

| Class | Description |
|-------|-------------|
| @WordDocumentParser.Models.Tables.TableData | Structured 2D table data (rows, columns, cells) |
| @WordDocumentParser.Models.Tables.TableRow | A row within a table |
| @WordDocumentParser.Models.Tables.TableCell | A cell within a row, with text content and nested table support |

### Enums

| Enum | Description |
|------|-------------|
| @WordDocumentParser.Core.ContentType | Node types: Document, Heading, Paragraph, Table, Image, List, ListItem, etc. |
| @WordDocumentParser.Models.ContentControls.ContentControlType | SDT control types |
