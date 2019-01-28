package net.auberson.wordwordcount;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import com.google.common.base.CharMatcher;
import com.google.common.base.Splitter;

public class Counter {

	final XWPFDocument doc;
	final XWPFStyles styleDefs;
	final Set<String> styles = new HashSet<String>();
	final List<String> outline = new ArrayList<String>();
	final List<Pattern> ignore = new ArrayList<Pattern>();
	final List<String> ignoreStyle = new ArrayList<String>();

	String startAfter, stopBefore;
	boolean debug;

	public static final Pattern CITATIONS = Pattern.compile("\\([^\\)]*?[12][0-9]{3}.*?\\)");

	public Counter(File file) throws IOException {
		FileInputStream is = new FileInputStream(file.getAbsolutePath());
		doc = new XWPFDocument(is);
		styleDefs = doc.getStyles();

		for (XWPFParagraph paragraph : doc.getParagraphs()) {
			String style = paragraph.getStyle();
			String text = paragraph.getText().trim();
			if (style != null) {
				styles.add(style);
			}
			if (isOutline(paragraph) && text.length() > 0) {
				outline.add(text);
			}
		}

	}

	private void ignore(Pattern pattern) {
		ignore.add(pattern);
	}

	private void ignoreStyle(String styleName) {
		ignoreStyle.add(styleName);
	}

	private void startAfter(String tocItem) {
		startAfter = tocItem;
	}

	private void stopBefore(String tocItem) {
		stopBefore = tocItem;
	}

	private void debug() {
		debug = true;
	}

	public Set<String> getUsedStyles() {
		return styles;
	}

	public List<String> getOutline() {
		return outline;
	}

	public List<Integer> getWordCount() {
		int count = 0;
		int ignoredCount = 0;
		int textboxCount = 0;

		boolean thisPartCounts = (startAfter == null ? true : false);

		paragraphs: for (XWPFParagraph paragraph : doc.getParagraphs()) {
			// Text in Textboxes don't count
			textboxCount += getWordCountInTextboxes(paragraph);

			// Check for starting and ending elements of the outline
			if (isOutline(paragraph)) {
				if (paragraph.getText().trim().equalsIgnoreCase(startAfter)) {
					thisPartCounts = true;
				}
				if (paragraph.getText().trim().equalsIgnoreCase(stopBefore)) {
					thisPartCounts = false;
				}
			}

			// Don't count if we haven't found the starting Outline element.
			if (!thisPartCounts) {
				ignoredCount += getWordCount(paragraph);
				if (debug) {
					debugOutput("IGNORED-OUTLINE:", paragraph);
				}
				continue paragraphs;
			}

			// Don't count styles that were specifically ignored.
			for (String styleName : ignoreStyle) {
				if (styleName.equalsIgnoreCase(paragraph.getStyle())) {
					ignoredCount += getWordCount(paragraph);
					if (debug) {
						debugOutput("IGNORED-" + paragraph.getStyle().toUpperCase() + ":", paragraph);
					}
					continue paragraphs;
				}
			}

			// Output any pattern that would be ignored
			if (debug) {
				debugOutput(ignore, paragraph);
			}

			// Count if we get until here
			count += getWordCount(paragraph);
		}
		return Arrays.asList(count, ignoredCount, textboxCount, count + ignoredCount + textboxCount);
	}

	private int getWordCount(XWPFParagraph paragraph) {
		String input = paragraph.getText();
		for (Pattern pattern : ignore) {
			input = pattern.matcher(input).replaceAll("");
		}

		List<String> words = Splitter.on(CharMatcher.whitespace()).splitToList(input);
		return words.size();
	}

	private int getWordCountInTextboxes(XWPFParagraph paragraph) {
		int totalCount = 0;
		StringBuilder debugText = new StringBuilder();

		XmlCursor cursor = paragraph.getCTP().newCursor();
		cursor.selectPath(
				"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");

		while (cursor.hasNextSelection()) {
			try {
				cursor.toNextSelection();
				XmlObject obj = cursor.getObject();
				CTR ctr;
				ctr = CTR.Factory.parse(obj.xmlText());
				XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
				String text = bufferrun.getText(0);
				if (text == null) {
					continue;
				}
				List<String> words = Splitter.on(CharMatcher.whitespace()).splitToList(text);
				totalCount += words.size();

				if (debug) {
					debugText.append(text);
				}
			} catch (XmlException e) {
				e.printStackTrace();
			}
		}

		if (debug) {
			debugOutput("IGNORED-TEXTBOX:", debugText.toString());
		}

		return totalCount;
	}

	private Integer getOutlineLevel(XWPFParagraph paragraph) {
		String styleName = paragraph.getStyle();
		if (styleName == null) {
			return null;
		}
		XWPFStyle style = styleDefs.getStyle(styleName);
		if (style == null) {
			System.err.println("Style not found: " + styleName);
			return null;
		}
		CTDecimalNumber outlineLvl = style.getCTStyle().getPPr().getOutlineLvl();
		if (outlineLvl == null) {
			return null;
		}
		return outlineLvl.getVal().intValue();
	}

	private boolean isOutline(XWPFParagraph paragraph) {
		return getOutlineLevel(paragraph) != null;
	}

	private void debugOutput(String prefix, XWPFParagraph paragraph) {
		debugOutput(prefix, paragraph.getText().trim());
	}

	private void debugOutput(String prefix, String text) {
		if (text.length() < 1) {
			return;
		}
		int sampleSize = 80 - prefix.length();
		if (text.length() < sampleSize) {
			System.out.println(prefix + " " + text.substring(0, text.length()));
		} else {
			System.out.println(prefix + " " + text.substring(0, sampleSize) + "...");
		}
	}

	private void debugOutput(List<Pattern> patterns, XWPFParagraph paragraph) {
		for (Pattern pattern : patterns) {
			Matcher matcher = pattern.matcher(paragraph.getText());
			while (matcher.find()) {
				debugOutput("IGNORED-PATTERN:", matcher.group());
			}
		}
	}

	/**
	 * Used for debugging patterns. Will dump any match found in the document for
	 * this pattern.
	 * 
	 * @param pattern
	 */
	public void findAll(Pattern pattern) {
		for (XWPFParagraph paragraph : doc.getParagraphs()) {
			Matcher matcher = pattern.matcher(paragraph.getText());
			while (matcher.find()) {
				System.out.println(matcher.group());
			}
		}
	}

	public static void main(String[] args) throws IOException {
		Counter counter = new Counter(new File(args[0]));
		counter.ignore(Counter.CITATIONS);
		counter.ignoreStyle("DXCCaptionText");
		counter.ignoreStyle("DXCPictures");
		counter.startAfter("Introduction");
		counter.stopBefore("Bibliography");

		System.out.println("Styles used:");
		System.out.println(counter.getUsedStyles());
		System.out.println();
		System.out.println("Document Outline:");
		System.out.println(counter.getOutline());
		System.out.println();

		System.out.println("Ignored sections:");
		counter.debug();
		List<Integer> wordCount = counter.getWordCount();
		System.out.println();

		System.out.println("Word Count [counted, ignored, textboxes, total]:");
		System.out.println(wordCount);
	}

}
