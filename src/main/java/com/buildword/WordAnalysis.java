package com.buildword;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class WordAnalysis {

	private final String defaultRegex = "\\$\\{[^{}]+\\}";
	private String tempFile;
	private String saveFile;

	@SuppressWarnings("resource")
	private void CopyFile() throws IOException {
		File tFile = new File(saveFile);
		//20190501 退出不删除文件
		//tFile.deleteOnExit();
		if (!tFile.getParentFile().exists()) {
			// 目标文件所在目录不存在
			tFile.getParentFile().mkdirs();
		}
		FileInputStream inStream = new FileInputStream(tempFile);
		FileOutputStream outStream = new FileOutputStream(tFile);
		FileChannel inC = inStream.getChannel();
		FileChannel outC = outStream.getChannel();
		int length = 2097152;
		while (true) {
			if (inC.position() == inC.size()) {
				inC.close();
				outC.close();
				tFile = null;
				inC = null;
				outC = null;
				break;
			}
			if ((inC.size() - inC.position()) < 20971520)
				length = (int) (inC.size() - inC.position());
			else
				length = 20971520;
			inC.transferTo(inC.position(), length, outC);
			inC.position(inC.position() + length);
		}

	};

	public WordAnalysis(String tempFile) {
		this.tempFile = tempFile;
		this.saveFile = tempFile;
	}

	public WordAnalysis(String tempFile, String saveFile) {
		this.tempFile = tempFile;
		this.saveFile = saveFile;
		// 复制模版文件到输出文件
		try {
			CopyFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 打开文档
	// 采用流的方式可以打开保存在统一个文集
	// opcpackage 必须保存为另外一个文件
	public XWPFDocument openDocument() throws IOException {
		XWPFDocument xdoc = null;
		InputStream is = null;
		is = new FileInputStream(saveFile);
		xdoc = new XWPFDocument(is);
		return xdoc;
	}

	// 关闭文档
	public void closeDocument(XWPFDocument document) {
		try {
			document.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 保存文档
	public void saveDocument(XWPFDocument document) {
		OutputStream os;
		try {
			os = new FileOutputStream(saveFile);
			if (os != null) {
				document.write(os);
				os.close();
			}
			closeDocument(document);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 复制Run
	private void CopyRun(XWPFRun target, XWPFRun source) {
		target.getCTR().setRPr(source.getCTR().getRPr());
		// 设置文本
		target.setText(source.text());
	}

	// 复制段落
	private void copyParagraph(XWPFParagraph target, XWPFParagraph source) {
		// 设置段落样式
		target.getCTP().setPPr(source.getCTP().getPPr());
		// 添加Run标签
		for (int pos = 0; pos < target.getRuns().size(); pos++) {
			target.removeRun(pos);
		}
		for (XWPFRun s : source.getRuns()) {
			XWPFRun targetrun = target.createRun();
			CopyRun(targetrun, s);
		}
	}

	// 复制单元格
	private void copyTableCell(XWPFTableCell target, XWPFTableCell source) {
		// 列属性
		target.getCTTc().setTcPr(source.getCTTc().getTcPr());
		// 删除目标 targetCell 所有单元格
		for (int pos = 0; pos < target.getParagraphs().size(); pos++) {
			target.removeParagraph(pos);
		}
		// 添加段落
		for (XWPFParagraph sp : source.getParagraphs()) {
			XWPFParagraph targetP = target.addParagraph();
			copyParagraph(targetP, sp);
		}
	}

	// 复制行
	private void CopytTableRow(XWPFTableRow target, XWPFTableRow source) {
		// 复制样式
		target.getCtRow().setTrPr(source.getCtRow().getTrPr());
		// 复制单元格
		for (int i = 0; i < target.getTableCells().size(); i++) {
			copyTableCell(target.getCell(i), source.getCell(i));
		}
	}

	// 获取表格中所有段落
	public List<XWPFParagraph> getTableParagraph(XWPFTable table) {
		List<XWPFParagraph> paras = new ArrayList<XWPFParagraph>();
		List<XWPFTableRow> rows = table.getRows();
		for (XWPFTableRow row : rows) {
			for (XWPFTableCell cell : row.getTableCells()) {
				for (XWPFParagraph p : cell.getParagraphs()) {
					// 去掉空白字符串
					if (p.getText() != null && p.getText().length() > 0) {
						paras.add(p);
					}
				}
			}
		}
		return paras;
	}

	// 返回为空 表示是普通表格，否则是个列表
	private String getTableListName(XWPFTable table, String regex) {
		if (regex == "")
			regex = defaultRegex;
		String tableName = "";
		XWPFTableRow firstRow = table.getRow(0);
		XWPFTableCell firstCell = firstRow.getCell(0);
		String cellText = firstCell.getText();
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(cellText);
		boolean find = matcher.find();
		while (find) {
			tableName = matcher.group();
			// 跳出循环
			find = false;
		}
		firstRow = null;
		firstCell = null;
		pattern = null;
		matcher = null;
		cellText = null;
		return tableName;

	}

	// 获取文档中所有的表格，不包含嵌套表格
	// listTable false 返回普通表格, true 返回列表表格
	public List<XWPFTable> getDocTables(XWPFDocument doc, boolean listTable,
			String regex) {
		List<XWPFTable> lstTables = new ArrayList<XWPFTable>();
		for (XWPFTable table : doc.getTables()) {
			String tbName = getTableListName(table, regex);
			if (listTable && tbName != "") {
				lstTables.add(table);
			}
			if (!listTable && (tbName == null || tbName.length() <= 0)) {
				lstTables.add(table);
			}
		}
		return lstTables;
	}

	// 向 taglist中添加新解析的段落信息
	private void setTagInfoList(List<TagInfo> list, XWPFParagraph p,
			String regex) {
		if (regex == "")
			regex = defaultRegex;
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(p.getText());
		int startPosition = 0;
		while (matcher.find(startPosition)) {
			String match = matcher.group();
			if (!list.contains(new TagInfo(match, ""))) {
				list.add(new TagInfo(match, ""));
			}
			startPosition = matcher.end();
		}
	}

	// 获取段落集合中所有文本
	public List<TagInfo> getWordTag(XWPFDocument doc, String regex) {
		List<TagInfo> tags = new ArrayList<TagInfo>();
		// 普通段落
		List<XWPFParagraph> pars = doc.getParagraphs();
		for (int i = 0; i < pars.size(); i++) {
			XWPFParagraph p = pars.get(i);
			setTagInfoList(tags, p, regex);
		}
		// Table中段落
		List<XWPFTable> commTables = getDocTables(doc, false, regex);
		for (XWPFTable table : commTables) {
			List<XWPFParagraph> tparags = getTableParagraph(table);
			for (int i = 0; i < tparags.size(); i++) {
				XWPFParagraph p = tparags.get(i);
				setTagInfoList(tags, p, regex);
			}
		}
		return tags;
	}

	// 获取Table列表中的配置信息
	public Map<String, List<List<TagInfo>>> getTableTag(XWPFDocument doc,
			String regex) {
		Map<String, List<List<TagInfo>>> mapList = new HashMap<String, List<List<TagInfo>>>();
		List<XWPFTable> lstTables = getDocTables(doc, true, regex);
		for (XWPFTable table : lstTables) {
			// 获取每个表格第一个单元格，以及最后一行
			String strTableName = getTableListName(table, regex);
			List<List<TagInfo>> list = new ArrayList<List<TagInfo>>();
			List<TagInfo> lstTag = new ArrayList<TagInfo>();
			int rowSize = table.getRows().size();
			XWPFTableRow lastRow = table.getRow(rowSize - 1);
			for (XWPFTableCell cell : lastRow.getTableCells()) {
				for (XWPFParagraph p : cell.getParagraphs()) {
					// 去掉空白字符串
					if (p.getText() != null && p.getText().length() > 0) {
						setTagInfoList(lstTag, p, regex);
					}
				}
			}
			list.add(lstTag);
			// 添加到数据集
			mapList.put(strTableName, list);
		}
		return mapList;
	}

	// 替换文本 已处理跨行的情况
	// 注意 文档中 不能出现类似$${\w+}的字符，由于searchText会一个字符一个字符做比价，找到第一个比配的开始计数
	public void ReplaceInParagraph(List<TagInfo> tagList, XWPFParagraph para,
			String regex) {
		if (regex == "")
			regex = defaultRegex;
		List<XWPFRun> runs = para.getRuns();
		for (TagInfo ti : tagList) {
			String find = ti.TagText;
			String replValue = ti.TagValue;
			TextSegement found = para.searchText(find,
					new PositionInParagraph());
			if (found != null) {
				// 判断查找内容是否在同一个Run标签中
				if (found.getBeginRun() == found.getEndRun()) {
					XWPFRun run = runs.get(found.getBeginRun());
					String runText = run.getText(run.getTextPosition());
					String replaced = runText.replace(find, replValue);
					run.setText(replaced, 0);
				} else {
					// 存在多个Run标签
					StringBuilder sb = new StringBuilder();
					for (int runPos = found.getBeginRun(); runPos <= found
							.getEndRun(); runPos++) {
						XWPFRun run = runs.get(runPos);
						sb.append(run.getText((run.getTextPosition())));
					}
					String connectedRuns = sb.toString();
					String replaced = connectedRuns.replace(find, replValue);
					XWPFRun firstRun = runs.get(found.getBeginRun());
					firstRun.setText(replaced, 0);
					// 删除后边的run标签
					for (int runPos = found.getBeginRun() + 1; runPos <= found
							.getEndRun(); runPos++) {
						// 清空其他标签内容
						XWPFRun partNext = runs.get(runPos);
						partNext.setText("", 0);
					}
				}
			}
		}
		// 完成第一遍查找,检测段落中的标签是否已经替换完 TODO 2016-06-14忘记当时处于什么考虑 加入这段代码
		// Pattern pattern = Pattern.compile(regex);
		// Matcher matcher = pattern.matcher(para.getText());
		// boolean find = matcher.find();
		// if (find) {
		// ReplaceInParagraph(tagList, para, regex);
		// find = false;
		// }
	}

	// 替换列表数据
	public void ReplaceInTable(List<List<TagInfo>> tagList, XWPFTable table,
			String regex) {
		int tempRowIndex = table.getRows().size() - 1;
		XWPFTableRow tempRow = table.getRow(tempRowIndex);
		for (List<TagInfo> lst : tagList) {
			table.createRow();
			XWPFTableRow newRow = table.getRow(table.getRows().size() - 1);
			CopytTableRow(newRow, tempRow);
			List<XWPFTableCell> nCells = newRow.getTableCells();
			for (int i = 0; i < nCells.size(); i++) {
				XWPFTableCell cell = newRow.getCell(i);
				for (XWPFParagraph p : cell.getParagraphs()) {
					if (p.getText() != null && p.getText().length() > 0) {
						ReplaceInParagraph(lst, p, regex);
					}
				}
			}
		}
		// 删除模版行
		table.removeRow(tempRowIndex);
	}

	// 替换所有tag
	public void ReplaceAllTag(XWPFDocument doc, List<TagInfo> formTagList,
			Map<String, List<List<TagInfo>>> tableTagList, String regex) {
		// 替换普通段落
		for (XWPFParagraph p : doc.getParagraphs()) {
			ReplaceInParagraph(formTagList, p, regex);
		}
		// 替换普通表格中段落
		List<XWPFTable> listCommTable = getDocTables(doc, false, regex);
		for (XWPFTable t : listCommTable) {
			List<XWPFParagraph> lstable = getTableParagraph(t);
			for (XWPFParagraph pt : lstable) {
				ReplaceInParagraph(formTagList, pt, regex);
			}
		}
		List<XWPFTable> listTable = getDocTables(doc, true, regex);
		for (XWPFTable table : listTable) {
			String tableName = getTableListName(table, regex);
			List<TagInfo> tableNameTags = new ArrayList<TagInfo>();
			tableNameTags.add(new TagInfo(tableName, ""));
			XWPFTableCell firstCell = table.getRow(0).getCell(0);
			List<XWPFParagraph> cellParas = firstCell.getParagraphs();
			for (XWPFParagraph pt : cellParas) {
				ReplaceInParagraph(tableNameTags, pt, regex);
			}
			List<List<TagInfo>> targetTableList = tableTagList.get(tableName);
			ReplaceInTable(targetTableList, table, regex);
		}
	}
}

