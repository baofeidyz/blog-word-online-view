package com.baofeidyz.blog;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.EditingLanguage;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.PdfLoadOptions;
import com.aspose.words.PdfPageMode;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.Revision;
import com.aspose.words.RevisionOptions;
import com.aspose.words.RevisionsView;
import com.aspose.words.ShowInBalloons;
import java.io.InputStream;

/**
 * 博客《技术选型-浏览器在线预览word》中提到的使用aspose将word转换为pdf代码示例.
 */
public class App {

    public static void main(String[] args) throws Exception {
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.CHINESE_PRC);
        try (InputStream inputStream = App.class.getClassLoader().getResourceAsStream("test.docx")) {
            Document document = new Document(inputStream, loadOptions);
            for (Revision revision : document.getRevisions()) {
                revision.setAuthor("author");
            }
            NodeCollection comments = document.getChildNodes(NodeType.COMMENT, true);

            for (Comment comment : (Iterable<Comment>) comments) {
                System.out.println("1= " + comment.getAuthor());
                comment.setInitial(comment.getAuthor());
            }
            document.getLayoutOptions().setShowParagraphMarks(true);
            document.getRevisions().get(0).getAuthor();
            document.setRevisionsView(RevisionsView.ORIGINAL);
            document.getLayoutOptions().setShowComments(true);
            document.setTrackRevisions(true);
            RevisionOptions revisionOptions = document.getLayoutOptions().getRevisionOptions();
            revisionOptions.setShowOriginalRevision(true);
            revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
            revisionOptions.setShowRevisionBars(true);
            revisionOptions.setShowRevisionMarks(true);
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            saveOptions.setPreserveFormFields(true);

            saveOptions.setAdditionalTextPositioning(true);
            saveOptions.setCreateNoteHyperlinks(true);
            saveOptions.setDisplayDocTitle(true);
            saveOptions.setExportDocumentStructure(true);
            saveOptions.setOpenHyperlinksInNewWindow(true);
            saveOptions.setPageMode(PdfPageMode.USE_OUTLINES);
            saveOptions.setPrettyFormat(true);
            document.save("target/result.pdf", saveOptions);
        }
    }

}
