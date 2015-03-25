package com.konidela.word2html.web;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xwpf.converter.core.IURIResolver;
import org.apache.poi.xwpf.converter.xhtml.DefaultContentHandlerFactory;
import org.apache.poi.xwpf.converter.xhtml.IContentHandlerFactory;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.konidela.word2html.converter.KWBXHTMLMapper;

/**
 * Servlet implementation class Word2Html
 */
@WebServlet("/Word2Html")
public class Word2Html extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public Word2Html() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		try {
			docx2Html(request, response);
		} catch (Exception e) {
			throw new ServletException(e);
		}
	}
	
	public void docx2Html(HttpServletRequest request, HttpServletResponse response) throws Exception {
		InputStream fis = null;
		try {
			 
			fis = new FileInputStream("E:\\Application Data\\MetricStream\\Attachments\\Sample data.docx");
			XWPFDocument document = new XWPFDocument(fis);
			final String imgUrl = "ImageLoader?imageId=";
			XHTMLOptions options = XHTMLOptions.create().URIResolver(new IURIResolver(){

				@Override
				public String resolve(String uri) {
					int ls = uri.lastIndexOf('/');
					if (ls >= 0)
						uri = uri.substring(ls+1);
					return imgUrl+uri;
				}});
			OutputStream out = response.getOutputStream();
			IContentHandlerFactory factory = options.getContentHandlerFactory();
			if (factory == null) {
				factory = DefaultContentHandlerFactory.INSTANCE;
			}
			options.setIgnoreStylesIfUnused(false);
			KWBXHTMLMapper mapper = new KWBXHTMLMapper(document, factory.create(out, null, options), options);
		    mapper.start();
		}
		catch(Exception ex) {
			throw ex;
		}
		finally {
			if(fis != null) {
				fis.close();
			}
		}
	}

}
