package com.konidela.word2html.converter;

import static org.apache.poi.xwpf.converter.core.utils.DxaUtil.emu2points;
import static org.apache.poi.xwpf.converter.xhtml.internal.XHTMLConstants.CLASS_ATTR;
import static org.apache.poi.xwpf.converter.xhtml.internal.XHTMLConstants.IMG_ELEMENT;
import static org.apache.poi.xwpf.converter.xhtml.internal.XHTMLConstants.SPAN_ELEMENT;
import static org.apache.poi.xwpf.converter.xhtml.internal.XHTMLConstants.SRC_ATTR;
import static org.apache.poi.xwpf.converter.xhtml.internal.XHTMLConstants.STYLE_ATTR;
import static org.apache.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.HEIGHT;
import static org.apache.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.WIDTH;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.converter.core.Color;
import org.apache.poi.xwpf.converter.core.ListItemContext;
import org.apache.poi.xwpf.converter.core.openxmlformats.styles.run.RunFontStyleStrikeValueProvider;
import org.apache.poi.xwpf.converter.core.openxmlformats.styles.run.RunTextHighlightingValueProvider;
import org.apache.poi.xwpf.converter.core.utils.ColorHelper;
import org.apache.poi.xwpf.converter.core.utils.StringUtils;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.converter.xhtml.internal.XHTMLMapper;
import org.apache.poi.xwpf.converter.xhtml.internal.styles.CSSStyle;
import org.apache.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants;
import org.apache.poi.xwpf.converter.xhtml.internal.utils.SAXHelper;
import org.apache.poi.xwpf.converter.xhtml.internal.utils.StringEscapeUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH.Enum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;

public class KWBXHTMLMapper extends XHTMLMapper {
	
	private ContentHandler contentHandler;

	private XWPFParagraph paragraph;
	
	private XWPFDocument document;
	
	private final List<String> allowedImageTypes = new ArrayList<String>();

	public KWBXHTMLMapper(XWPFDocument document, ContentHandler contentHandler, XHTMLOptions options) throws Exception {
		super(document, contentHandler, options);
		this.contentHandler = contentHandler;
		this.document = document;
		allowedImageTypes.add("jpg");
		allowedImageTypes.add("png");
		allowedImageTypes.add("jpeg");
		allowedImageTypes.add("gif"); //add more as needed
	}
	
	@Override
	protected Object startVisitParagraph(XWPFParagraph paragraph, ListItemContext itemContext, Object parentContainer)
			throws Exception {
		Object result = super.startVisitParagraph(paragraph, itemContext, parentContainer);
		if (itemContext != null) {
			AttributesImpl attributes = createClassAttribute(paragraph.getStyleID());

			CTPPr pPr = paragraph.getCTP().getPPr();
			CSSStyle cssStyle = getStylesDocument().createCSSStyle(pPr);
			attributes = createStyleAttribute(cssStyle, attributes);
			startElement(SPAN_ELEMENT, attributes);
			String text = itemContext.getText();
			if (StringUtils.isNotEmpty(text)) {
				text = replaceNonUnicodeChars(text);
				text = text + " ";
				SAXHelper.characters(contentHandler, StringEscapeUtils.escapeHtml(text));
			}
			endElement(SPAN_ELEMENT);
		}
		return result;
	}
	
	@Override
	protected void visitRun(XWPFRun run, boolean pageNumber, String url, Object paragraphContainer) throws Exception {
		boolean isUrl = url != null;
		if (isUrl) {
			AttributesImpl hyperlinkAttributes = new AttributesImpl();
			SAXHelper.addAttrValue(hyperlinkAttributes, "href", url);
			SAXHelper.addAttrValue(hyperlinkAttributes, "target", "_blank");
			startElement("a", hyperlinkAttributes);
			url = null;
		}
		if (run.getFontFamily() == null) {
			run.setFontFamily(getStylesDocument().getFontFamilyAscii(run));
		}
		if (run.getFontSize() <= 0) {
			run.setFontSize(getStylesDocument().getFontSize(run).intValue());
		}
		CTRPr rPr = run.getCTR().getRPr();
		if (run.getParent() instanceof XWPFParagraph) {
			paragraph = (XWPFParagraph) run.getParent();
		}
		if (rPr != null
				&& (rPr.getHighlight() != null || rPr.getStrike() != null || rPr.getDstrike() != null || rPr
						.getVertAlign() != null) && paragraph != null) {
			StringBuilder text = new StringBuilder();
			XmlCursor c = run.getCTR().newCursor();
			c.selectPath("./*");
			while (c.toNextSelection()) {
				XmlObject o = c.getObject();
				if (o instanceof CTText) {
					if (!"w:instrText".equals(o.getDomNode().getNodeName())) {
						text.append(((CTText) o).getStringValue());
					}
				}
			}
			// 1) create attributes

			// 1.1) Create "class" attributes.
			AttributesImpl runAttributes = createClassAttribute(paragraph.getStyleID());
			boolean isSuper = false;
			boolean isSub = false;

			// 1.2) Create "style" attributes.
			CSSStyle cssStyle = getStylesDocument().createCSSStyle(rPr);
			if (cssStyle != null) {
				Color color = RunTextHighlightingValueProvider.INSTANCE.getValue(rPr, getStylesDocument());
				if (color != null) {
					cssStyle.addProperty(CSSStylePropertyConstants.BACKGROUND_COLOR, ColorHelper.toHexString(color));
				}
				if (Boolean.TRUE.equals(RunFontStyleStrikeValueProvider.INSTANCE.getValue(rPr, getStylesDocument()))
						|| rPr.getDstrike() != null) {
					cssStyle.addProperty("text-decoration", "line-through");
				}
				if (rPr.getVertAlign() != null) {
					int align = rPr.getVertAlign().getVal().intValue();
					if (STVerticalAlignRun.INT_SUPERSCRIPT == align) {

						isSuper = true;
					} else if (STVerticalAlignRun.INT_SUBSCRIPT == align) {

						isSub = true;
					}
				}
			}
			runAttributes = createStyleAttribute(cssStyle, runAttributes);
			if (runAttributes != null) {
				startElement(SPAN_ELEMENT, runAttributes);
				if (isSuper || isSub) {
					startElement(isSuper ? "SUP" : "SUB", null);
				}
			}
			String txt = text.toString();
			if (StringUtils.isNotEmpty(txt)) {
				// Escape with HTML characters
				characters(StringEscapeUtils.escapeHtml(txt));
			}
			if (runAttributes != null) {
				if (isSuper || isSub) {
					endElement(isSuper ? "SUP" : "SUB");
				}
				endElement(SPAN_ELEMENT);
			}

			if (isUrl) {
				characters(" ");
				endElement("a");
			}
			return;
		}
		super.visitRun(run, pageNumber, url, paragraphContainer);
		if (isUrl) {
			characters(" ");
			endElement("a");
		}
		paragraph = null;
	}
	
	@Override
	protected void visitTabs(CTTabs o, Object paragraphContainer) throws Exception {
		if (paragraph != null && o == null) {
			startElement(SPAN_ELEMENT, null);
			characters("&nbsp;&nbsp;");//add no of space as needed for each tab
			endElement(SPAN_ELEMENT);
			return;
		}
		super.visitTabs(o, paragraphContainer);
	}
	
	@Override
	protected void visitPicture(CTPicture picture, Float offsetX, Enum relativeFromH, Float offsetY,
			org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV.Enum relativeFromV,
			org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STWrapText.Enum wrapText,
			Object parentContainer) throws Exception {

		XWPFPictureData pictureData = getPictureData(picture);
		if (pictureData != null) {
			super.visitPicture(picture, offsetX, relativeFromH, offsetY, relativeFromV, wrapText, parentContainer);
		}

		else {
			// external link images inserted
			String link = picture.getBlipFill().getBlip().getLink();
			PackageRelationship rel = document.getPackagePart().getRelationships().getRelationshipByID(link);
			if (rel != null) {
				String src = rel.getTargetURI().toString();
				//if static image, we can control what image types to render
				if(src.lastIndexOf(".") > 0) {
					String extension = src.substring(src.lastIndexOf(".") + 1);
					if(allowedImageTypes.contains(extension) == false) {
						return;
					}
				}
				AttributesImpl attributes = SAXHelper.addAttrValue(null, SRC_ATTR, src);

				CTPositiveSize2D ext = picture.getSpPr().getXfrm().getExt();

				// img/@width
				float width = emu2points(ext.getCx());
				attributes = SAXHelper.addAttrValue(attributes, WIDTH, getStylesDocument().getValueAsPoint(width));

				// img/@height
				float height = emu2points(ext.getCy());
				attributes = SAXHelper.addAttrValue(attributes, HEIGHT, getStylesDocument().getValueAsPoint(height));

				if (attributes != null) {
					SAXHelper.startElement(contentHandler, IMG_ELEMENT, attributes);
					SAXHelper.endElement(contentHandler, IMG_ELEMENT);
				}
			}
		}
	}
	
	private AttributesImpl createClassAttribute(String styleID) {
		String classNames = getStylesDocument().getClassNames(styleID);
		if (StringUtils.isNotEmpty(classNames)) {
			return SAXHelper.addAttrValue(null, CLASS_ATTR, classNames);
		}
		return null;
	}
	
	private AttributesImpl createStyleAttribute(CSSStyle cssStyle, AttributesImpl attributes) {
		if (cssStyle != null) {
			String inlineStyles = cssStyle.getInlineStyles();
			if (StringUtils.isNotEmpty(inlineStyles)) {
				attributes = SAXHelper.addAttrValue(attributes, STYLE_ATTR, inlineStyles);
			}
		}
		return attributes;
	}
	
	private void startElement(String name, Attributes attributes) throws SAXException {
		SAXHelper.startElement(contentHandler, name, attributes);
	}

	private void endElement(String name) throws SAXException {
		SAXHelper.endElement(contentHandler, name);
	}

	private void characters(String content) throws SAXException {
		SAXHelper.characters(contentHandler, content);
	}
	
	public static String replaceNonUnicodeChars(String text) {
		StringBuilder newString = new StringBuilder(text.length());
		for (int offset = 0; offset < text.length();) {
			int codePoint = text.codePointAt(offset);
			offset += Character.charCount(codePoint);

			// Replace invisible control characters and unused code points
			switch (Character.getType(codePoint)) {
				case Character.CONTROL:
				case Character.FORMAT:
				case Character.PRIVATE_USE:
				case Character.SURROGATE:
				case Character.UNASSIGNED: {
					newString.append('\u2022');
					break;
				}
				default: {
					newString.append(Character.toChars(codePoint));
					break;
				}
			}
		}
		return newString.toString();
	}
}
