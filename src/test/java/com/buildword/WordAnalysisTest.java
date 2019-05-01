package com.buildword;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import junit.framework.TestCase;

//author herbert QQ 464884492
public class WordAnalysisTest extends TestCase {
	public void testBuildWord() {

		File f = new File("1.txt");
		System.out.println(f.getAbsolutePath());
		String tempFile = "src/word.docx";
		SimpleDateFormat fmt = new SimpleDateFormat("yyyyMMddHHmmss");
		String outFile = "src/" + fmt.format(new Date()) + ".docx";

		WordAnalysis wa = new WordAnalysis(tempFile, outFile);
		try {
			XWPFDocument d = wa.openDocument();
			// 获取普通标签
			List<TagInfo> lst = wa.getWordTag(d, "");
			for (TagInfo i : lst) {
				System.out.println(i.TagText);
				i.TagValue = "hhh";
			}
			// 获取列表标签
			Map<String, List<List<TagInfo>>> maplist = wa.getTableTag(d, "");

			for (Map.Entry<String, List<List<TagInfo>>> entry : maplist.entrySet()) {
				// 输出表格名称
				System.out.println(entry.getKey());
				// 输出列
				for (TagInfo ti : entry.getValue().get(0)) {
					System.out.print(ti.TagText + "---");
				}
				System.out.println();
			}
			// 替换普通段落
			for (XWPFParagraph p : d.getParagraphs()) {
				wa.ReplaceInParagraph(lst, p, "");
			}
			// 替换普通表格
			List<XWPFTable> tbls = wa.getDocTables(d, false, "");
			for (XWPFTable table : tbls) {
				for (XWPFParagraph p : wa.getTableParagraph(table)) {
					wa.ReplaceInParagraph(lst, p, "");
				}
			}
			// 替换标准列表
			List<XWPFTable> listTable = wa.getDocTables(d, true, "");
			for (XWPFTable table : listTable) {
				List<List<TagInfo>> tagList = new ArrayList<List<TagInfo>>();
				List<TagInfo> row = new ArrayList<TagInfo>();
				row.add(new TagInfo("${COL1}", "测试1-1"));
				row.add(new TagInfo("${COL2}", "测试1-2"));

				List<TagInfo> row2 = new ArrayList<TagInfo>();
				row2.add(new TagInfo("${COL1}", "测试2-1"));
				row2.add(new TagInfo("${COL2}", "测试2-2"));

				List<TagInfo> row3 = new ArrayList<TagInfo>();
				row3.add(new TagInfo("${COL1}", "测试3-1"));
				row3.add(new TagInfo("${COL2}", "测试3-2"));

				tagList.add(row);
				tagList.add(row2);
				tagList.add(row3);
				wa.ReplaceInTable(tagList, table, "");
			}
			wa.saveDocument(d);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
