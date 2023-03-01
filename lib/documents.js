/* eslint-disable no-console */
var _ = require("underscore");

var types = exports.types = {
    document: "document",
    paragraph: "paragraph",
    run: "run",
    text: "text",
    tab: "tab",
    hyperlink: "hyperlink",
    noteReference: "noteReference",
    image: "image",
    chart: "chart",
    customDocDesc: 'customDocDesc',
    note: "note",
    commentReference: "commentReference",
    comment: "comment",
    table: "table",
    tableRow: "tableRow",
    tableCell: "tableCell",
    "break": "break",
    bookmarkStart: "bookmarkStart"
};

function Document(children, options) {
    options = options || {};
    return {
        type: types.document,
        children: children,
        notes: options.notes || new Notes({}),
        comments: options.comments || []
    };
}

function Paragraph(children, properties) {
    properties = properties || {};
    var paragraphProperties = {
        type: types.paragraph,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null
    };

    return Object.assign({}, properties, paragraphProperties);
}

function Run(children, properties) {
    properties = properties || {};
    return {
        type: types.run,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null,
        bold: properties.bold,
        'bold-complex-script': properties['bold-complex-script'],
        underline: properties.underline,
        italics: properties.italics,
        'italics-complex-script': properties['italics-complex-script'],
        'all-caps': properties['all-caps'],
        'small-caps': properties['small-caps'],
        'vertical-alignment': properties['vertical-alignment'],
        font: properties.font || null,
        fontSize: properties.fontSize || null,
        // 文本新增背景色和字体颜色
        bgColor: properties.bgColor || null,
        color: properties.color || null,
        spacing: properties.spacing || null,  // 新增缩进和行高
        commentId: properties.commentId || null, // 新增标记id位置
        effect: properties.effect || null,
        kern: properties.kern || null,
        size: properties.size || null,
        'size-complex-script': properties['size-complex-script'] || null,
        strike: properties.strike || null,
        'double-strike': properties['double-strike'] || null,
        highlight: properties.highlight || null,
        'highlight-complex-script': properties['highlight-complex-script'] || null,
        'character-spacing': properties['character-spacing'] || null,
        shading: properties.shading || null,
        emboss: properties.emboss || null,
        imprint: properties.imprint || null,
        language: properties.language || null,
        border: properties.border || null,
        'snap-to-grid': properties['snap-to-grid'] || null,
        vanish: properties.vanish || null,
        'spec-vanish': properties['spec-vanish'] || null,
        scale: properties.scale || null,
        math: properties.math || null
    };
}

var verticalAlignment = {
    baseline: "baseline",
    superscript: "superscript",
    subscript: "subscript"
};

function Text(value) {
    return {
        type: types.text,
        value: value
    };
}

function Tab() {
    return {
        type: types.tab
    };
}

function Hyperlink(children, options) {
    return {
        type: types.hyperlink,
        children: children,
        href: options.href,
        anchor: options.anchor,
        targetFrame: options.targetFrame
    };
}

function NoteReference(options) {
    return {
        type: types.noteReference,
        noteType: options.noteType,
        noteId: options.noteId
    };
}

function Notes(notes) {
    this._notes = _.indexBy(notes, function(note) {
        return noteKey(note.noteType, note.noteId);
    });
}

Notes.prototype.resolve = function(reference) {
    return this.findNoteByKey(noteKey(reference.noteType, reference.noteId));
};

Notes.prototype.findNoteByKey = function(key) {
    return this._notes[key] || null;
};

function Note(options) {
    return {
        type: types.note,
        noteType: options.noteType,
        noteId: options.noteId,
        body: options.body
    };
}

function commentReference(options) {
    return {
        type: types.commentReference,
        commentId: options.commentId
    };
}

function comment(options) {
    return {
        type: types.comment,
        commentId: options.commentId,
        body: options.body,
        authorName: options.authorName,
        authorInitials: options.authorInitials
    };
}

function noteKey(noteType, id) {
    return noteType + "-" + id;
}

function Image(options) {
    return {
        type: types.image,
        read: options.readImage,
        altText: options.altText,
        contentType: options.contentType
    };
}

function Chart(options) {
    return {
        type: types.chart,
        id: options.id,
        commentId: options.commentId || null, // 新增标记id位置
        altText: options.altText
    };
}

function CustomDocDesc(options) {
    options = options || {};
    options.descType = 'customDocDesc';
    return options;
}

function Table(children, properties) {
    properties = properties || {};
    return {
        type: types.table,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null,
        float: properties.float || null,
        overlap: properties.overlap || null,
        width: properties.width || null,
        indent: properties.indent || null,
        layout: properties.layout || null,
        cellMargin: properties.cellMargin || null,
        columnWidths: properties.columnWidths || null,
        visuallyRightToLeft: properties.visuallyRightToLeft || null,
        borders: properties.borders || null
    };
}

function TableRow(children, options) {
    options = options || {};
    return {
        type: types.tableRow,
        children: children,
        height: options.height || null,
        isHeader: options.isHeader || false
    };
}

function TableCell(children, options) {
    options = options || {};
    return {
        type: types.tableCell,
        children: children,
        colSpan: options.colSpan == null ? 1 : options.colSpan,
        rowSpan: options.rowSpan == null ? 1 : options.rowSpan,
        borders: options.borders || null,
        shading: options.shading || null,
        width: options.width || null,
        verticalMerge: options.verticalMerge || null,
        verticalAlign: options.verticalAlign || null,
        margins: options.margins || null,
        textDirection: options.textDirection || null
    };
}

function Break(breakType) {
    return {
        type: types["break"],
        breakType: breakType
    };
}

function BookmarkStart(options) {
    return {
        type: types.bookmarkStart,
        name: options.name
    };
}

exports.document = exports.Document = Document;
exports.paragraph = exports.Paragraph = Paragraph;
exports.run = exports.Run = Run;
exports.Text = Text;
exports.tab = exports.Tab = Tab;
exports.Hyperlink = Hyperlink;
exports.noteReference = exports.NoteReference = NoteReference;
exports.Notes = Notes;
exports.Note = Note;
exports.commentReference = commentReference;
exports.comment = comment;
exports.Image = Image;
exports.Chart = Chart;
exports.CustomDocDesc = CustomDocDesc;
exports.Table = Table;
exports.TableRow = TableRow;
exports.TableCell = TableCell;
exports.lineBreak = Break("line");
exports.pageBreak = Break("page");
exports.columnBreak = Break("column");
exports.BookmarkStart = BookmarkStart;

exports.verticalAlignment = verticalAlignment;
