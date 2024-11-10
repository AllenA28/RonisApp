package dataimporter.implementation.utils;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Iterator;
import java.util.LinkedList;
import java.util.Queue;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

public class ExtendedXSSFSheetXMLHandler extends DefaultHandler {
    private static final Logger LOG = LogManager.getLogger(ExtendedXSSFSheetXMLHandler.class);
    /**
     * Table with the styles used for formatting
     */
    private final Styles stylesTable;
    /**
     * Table with cell comments
     */
    private final Comments comments;
    /**
     * Read only access to the shared strings table, for looking
     * up (most) string cell's contents
     */
    private final SharedStrings sharedStringsTable;
    /**
     * Where our text is going
     */
    private final ExtendedXSSFSheetXMLHandler.SheetContentsHandler output;
    private final DataFormatter formatter;
    private final boolean formulasNotResults;
    // Gathers characters as they are seen.
    private final StringBuilder value = new StringBuilder(64);
    private final StringBuilder formula = new StringBuilder(64);
    private final StringBuilder headerFooter = new StringBuilder(64);
    // Set when V start element is seen
    private boolean vIsOpen;
    // Set when F start element is seen
    private boolean fIsOpen;
    // Set when an Inline String "is" is seen
    private boolean isIsOpen;
    // Set when a header/footer element is seen
    private boolean hfIsOpen;
    // Set when cell start element is seen
    // used when cell close element is seen.
    private ExtendedXSSFSheetXMLHandler.xssfDataType nextDataType;
    private CellType cellType;
    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private int rowNum;
    private int nextRowNum;      // some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
    private String cellRef;
    private Queue<CellAddress> commentCellRefs;

    /**
     * Accepts objects needed while parsing.
     *
     * @param styles  Table of styles
     * @param strings Table of shared strings
     */
    public ExtendedXSSFSheetXMLHandler(
            Styles styles,
            Comments comments,
            SharedStrings strings,
            ExtendedXSSFSheetXMLHandler.SheetContentsHandler sheetContentsHandler,
            DataFormatter dataFormatter,
            boolean formulasNotResults) {
        this.stylesTable = styles;
        this.comments = comments;
        this.sharedStringsTable = strings;
        this.output = sheetContentsHandler;
        this.formulasNotResults = formulasNotResults;
        this.nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.NUMBER;
        this.cellType = CellType.NUMERIC;
        this.formatter = dataFormatter;
        init(comments);
    }

    /**
     * Accepts objects needed while parsing.
     *
     * @param styles  Table of styles
     * @param strings Table of shared strings
     */
    public ExtendedXSSFSheetXMLHandler(
            Styles styles,
            SharedStrings strings,
            ExtendedXSSFSheetXMLHandler.SheetContentsHandler sheetContentsHandler,
            DataFormatter dataFormatter,
            boolean formulasNotResults) {
        this(styles, null, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
    }

    private void init(Comments commentsTable) {
        if (commentsTable != null) {
            commentCellRefs = new LinkedList<>();
            for (Iterator<CellAddress> iter = commentsTable.getCellAddresses(); iter.hasNext(); ) {
                commentCellRefs.add(iter.next());
            }
        }
    }

    private boolean isTextTag(String name) {
        if ("v".equals(name)) {
            // Easy, normal v text tag
            return true;
        }
        if ("inlineStr".equals(name)) {
            // Easy inline string
            return true;
        }
        // Inline string <is><t>...</t></is> pair
        return "t".equals(name) && isIsOpen;
        // It isn't a text tag
    }

    @Override
    @SuppressWarnings("unused")
    public void startElement(String uri, String localName, String qName,
                             Attributes attributes) throws SAXException {

        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if (isTextTag(localName)) {
            vIsOpen = true;
            // Clear contents cache
            if (!isIsOpen) {
                value.setLength(0);
            }
        } else if ("is".equals(localName)) {
            // Inline string outer tag
            isIsOpen = true;
        } else if ("f".equals(localName)) {
            // Clear contents cache
            formula.setLength(0);

            // Mark us as being a formula if not already
            if (this.nextDataType == ExtendedXSSFSheetXMLHandler.xssfDataType.NUMBER) {
                this.nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.FORMULA;
                this.cellType = CellType.FORMULA;
            }

            // Decide where to get the formula string from
            String type = attributes.getValue("t");
            if (type != null && type.equals("shared")) {
                // Is it the one that defines the shared, or uses it?
                String ref = attributes.getValue("ref");
                String si = attributes.getValue("si");

                if (ref != null) {
                    // This one defines it
                    fIsOpen = true;
                } else {
                    // This one uses a shared formula
                    if (formulasNotResults) {
                        LOG.atWarn().log("shared formulas not yet supported!");
                    }
                }
            } else {
                fIsOpen = true;
            }
        } else if ("oddHeader".equals(localName) || "evenHeader".equals(localName) ||
                "firstHeader".equals(localName) || "firstFooter".equals(localName) ||
                "oddFooter".equals(localName) || "evenFooter".equals(localName)) {
            hfIsOpen = true;
            // Clear contents cache
            headerFooter.setLength(0);
        } else if ("row".equals(localName)) {
            String rowNumStr = attributes.getValue("r");
            if (rowNumStr != null) {
                rowNum = Integer.parseInt(rowNumStr) - 1;
            } else {
                rowNum = nextRowNum;
            }
            output.startRow(rowNum);
        }
        // c => cell
        else if ("c".equals(localName)) {
            // Set up defaults.
            this.formula.setLength(0);
            this.nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.NUMBER;
            this.cellType = CellType.NUMERIC;
            this.formatIndex = -1;
            this.formatString = null;
            cellRef = attributes.getValue("r");
            String cellTypeValue = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellTypeValue)) {
                nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.BOOLEAN;
                this.cellType = CellType.BOOLEAN;
            } else if ("e".equals(cellTypeValue)) {
                nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.ERROR;
                this.cellType = CellType.ERROR;
            } else if ("inlineStr".equals(cellTypeValue)) {
                nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.INLINE_STRING;
                this.cellType = CellType.STRING;
            } else if ("s".equals(cellTypeValue)) {
                nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.SST_STRING;
                this.cellType = CellType.STRING;
            } else if ("str".equals(cellTypeValue)) {
                nextDataType = ExtendedXSSFSheetXMLHandler.xssfDataType.FORMULA;
                this.cellType = CellType.FORMULA;
            } else {
                // Number, but almost certainly with a special style or format
                XSSFCellStyle style = null;
                if (stylesTable != null) {
                    if (cellStyleStr != null) {
                        var styleIndex = Integer.parseInt(cellStyleStr);
                        style = stylesTable.getStyleAt(styleIndex);
                    } else if (stylesTable.getNumCellStyles() > 0) {
                        style = stylesTable.getStyleAt(0);
                    }
                }
                if (style != null) {
                    this.formatIndex = style.getDataFormat();
                    this.formatString = style.getDataFormatString();
                    if (this.formatString == null)
                        this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName)
            throws SAXException {

        if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        // v => contents of a cell
        if (isTextTag(localName)) {
            vIsOpen = false;

            if (!isIsOpen) {
                outputCell();
                value.setLength(0);
            }
        } else if ("f".equals(localName)) {
            fIsOpen = false;
        } else if ("is".equals(localName)) {
            isIsOpen = false;
            outputCell();
            value.setLength(0);
        } else if ("row".equals(localName)) {
            // Handle any "missing" cells which had comments attached
            checkForEmptyCellComments(ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW);

            // Finish up the row
            output.endRow(rowNum);

            // some sheets do not have rowNum set in the XML, Excel can read them so we should try to read them as well
            nextRowNum = rowNum + 1;
        } else if ("sheetData".equals(localName)) {
            // Handle any "missing" cells which had comments attached
            checkForEmptyCellComments(ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_SHEET_DATA);

            // indicate that this sheet is now done
            output.endSheet();
        } else if ("oddHeader".equals(localName) || "evenHeader".equals(localName) ||
                "firstHeader".equals(localName)) {
            hfIsOpen = false;
            output.headerFooter(headerFooter.toString(), true, localName);
        } else if ("oddFooter".equals(localName) || "evenFooter".equals(localName) ||
                "firstFooter".equals(localName)) {
            hfIsOpen = false;
            output.headerFooter(headerFooter.toString(), false, localName);
        }
    }

    /**
     * Captures characters only if a suitable element is open.
     * Originally was just "v"; extended for inlineStr also.
     */
    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
        if (fIsOpen) {
            formula.append(ch, start, length);
        }
        if (hfIsOpen) {
            headerFooter.append(ch, start, length);
        }
    }

    private void outputCell() {
        String thisStr = null;

        // Process the value contents as required, now we have it all
        if (formulasNotResults && formula.length() > 0) {
            thisStr = formula.toString();
        } else {
            switch (nextDataType) {
                case BOOLEAN:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;

                case ERROR:
                    thisStr = "ERROR:" + value;
                    break;

                case FORMULA:
                    if (formulasNotResults) {
                        thisStr = formula.toString();
                    } else {
                        var fv = value.toString();

                        if (this.formatString != null) {
                            try {
                                // Try to use the value as a formattable number
                                var d = Double.parseDouble(fv);
                                thisStr = formatter.formatRawCellContents(d, this.formatIndex, this.formatString);
                            } catch (NumberFormatException e) {
                                // Formula is a String result not a Numeric one
                                thisStr = fv;
                            }
                        } else {
                            // No formatting applied, just do raw value in all cases
                            thisStr = fv;
                        }
                    }
                    break;

                case INLINE_STRING:
                    var rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    break;

                case SST_STRING:
                    var sstIndex = value.toString();
                    if (sstIndex.length() > 0) {
                        try {
                            var idx = Integer.parseInt(sstIndex);
                            RichTextString rtss = sharedStringsTable.getItemAt(idx);
                            thisStr = rtss.toString();
                        } catch (NumberFormatException ex) {
                            LOG.atError().withThrowable(ex).log("Failed to parse SST index '{}'", sstIndex);
                        }
                    }
                    break;

                case NUMBER:
                    var n = value.toString();
                    if (this.formatString != null && n.length() > 0)
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
                    else
                        thisStr = n;
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    break;
            }
        }

        // Do we have a comment for this cell?
        checkForEmptyCellComments(ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.CELL);
        XSSFComment comment = comments != null ? comments.findCellComment(new CellAddress(cellRef)) : null;

        // Output
        output.cell(cellRef, thisStr, value.toString(), this.cellType, this.formatString, comment, this.formatIndex);
    }

    /**
     * Do a check for, and output, comments in otherwise empty cells.
     */
    private void checkForEmptyCellComments(ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType type) {
        if (commentCellRefs != null && !commentCellRefs.isEmpty()) {
            // If we've reached the end of the sheet data, output any
            //  comments we haven't yet already handled
            if (type == ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_SHEET_DATA) {
                while (!commentCellRefs.isEmpty()) {
                    outputEmptyCellComment(commentCellRefs.remove());
                }
                return;
            }

            // At the end of a row, handle any comments for "missing" rows before us
            if (this.cellRef == null) {
                if (type == ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW) {
                    while (!commentCellRefs.isEmpty()) {
                        if (commentCellRefs.peek().getRow() == rowNum) {
                            outputEmptyCellComment(commentCellRefs.remove());
                        } else {
                            return;
                        }
                    }
                    return;
                } else {
                    throw new IllegalStateException("Cell ref should be null only if there are only empty cells in the row; rowNum: " + rowNum);
                }
            }

            CellAddress nextCommentCellRef;
            do {
                var cellAddress = new CellAddress(this.cellRef);
                var peekCellRef = commentCellRefs.peek();
                if (type == ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.CELL && cellAddress.equals(peekCellRef)) {
                    // remove the comment cell ref from the list if we're about to handle it alongside the cell content
                    commentCellRefs.remove();
                    return;
                } else {
                    // fill in any gaps if there are empty cells with comment mixed in with non-empty cells
                    int comparison = peekCellRef.compareTo(cellAddress);
                    if (comparison > 0 && type == ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.END_OF_ROW && peekCellRef.getRow() <= rowNum) {
                        nextCommentCellRef = commentCellRefs.remove();
                        outputEmptyCellComment(nextCommentCellRef);
                    } else if (comparison < 0 && type == ExtendedXSSFSheetXMLHandler.EmptyCellCommentsCheckType.CELL && peekCellRef.getRow() <= rowNum) {
                        nextCommentCellRef = commentCellRefs.remove();
                        outputEmptyCellComment(nextCommentCellRef);
                    } else {
                        nextCommentCellRef = null;
                    }
                }
            } while (nextCommentCellRef != null && !commentCellRefs.isEmpty());
        }
    }

    /**
     * Output an empty-cell comment.
     */
    private void outputEmptyCellComment(CellAddress cellRef) {
        XSSFComment comment = comments.findCellComment(cellRef);
        output.cell(cellRef.formatAsString(), null, null, CellType.BLANK, null, comment, this.formatIndex);
    }


    /**
     * These are the different kinds of cells we support.
     * We keep track of the current one between
     * the start and end.
     */
    enum xssfDataType {
        BOOLEAN,
        ERROR,
        FORMULA,
        INLINE_STRING,
        SST_STRING,
        NUMBER,
    }

    private enum EmptyCellCommentsCheckType {
        CELL,
        END_OF_ROW,
        END_OF_SHEET_DATA
    }

    /**
     * This interface allows to provide callbacks when reading
     * a sheet in streaming mode.
     * <p>
     * The XSLX file is usually read via {@link XSSFReader}.
     * <p>
     * By implementing the methods, you can process arbitrarily
     * large files without exhausting main memory.
     */
    public interface SheetContentsHandler {
        /**
         * A row with the (zero based) row number has started
         */
        void startRow(int rowNum);

        /**
         * A row with the (zero based) row number has ended
         */
        void endRow(int rowNum) throws SAXException;

        /**
         * A cell, with the given formatted value (may be null),
         * and possibly a comment (may be null), was encountered.
         * <p>
         * Sheets that have missing or empty cells may result in
         * sparse calls to <code>cell</code>. See the code in
         * <code>poi-examples/src/main/java/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java</code>
         * for an example of how to handle this scenario.
         */
        void cell(String cellReference, String formattedValue, String rawValue, CellType cellType, String formatString, XSSFComment comment, short dataFormat);

        /**
         * A header or footer has been encountered
         */
        default void headerFooter(String text, boolean isHeader, String tagName) {
        }

        /**
         * Signal that the end of a sheet was been reached
         */
        default void endSheet() {
        }
    }
}