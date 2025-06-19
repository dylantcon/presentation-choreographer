package com.presentationchoreographer.xml.writers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.*;
import java.io.*;
import java.util.*;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.utils.XMLConstants;
import com.presentationchoreographer.exceptions.*;

/**
 * Core XML writer for injecting shapes, animations, and content into PowerPoint slides
 * Handles DOM manipulation and ensures OOXML compliance
 */
public class SlideXMLWriter {

  private final Document document;
  private final XPath xpath;
  private final Element shapeTree;
  private int nextAvailableSpid;

  // Namespace prefixes for creating new elements
  private static final String PRESENTATION_PREFIX = "p";
  private static final String DRAWING_PREFIX = "a";

  public SlideXMLWriter(Document document) throws XMLParsingException {
    this.document = document;

    try {
      XPathFactory xpathFactory = XPathFactory.newInstance();
      this.xpath = xpathFactory.newXPath();
      this.xpath.setNamespaceContext(new PowerPointNamespaceContext());

      // Find the shape tree where we'll inject new shapes
      this.shapeTree = (Element) xpath.evaluate("//p:spTree", document, XPathConstants.NODE);
      if (shapeTree == null) {
        throw new XMLParsingException("No shape tree found in slide document");
      }

      // Calculate next available SPID
      this.nextAvailableSpid = calculateNextSpid();

    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to initialize XML writer", e);
    }
  }

  /**
   * Inject a basic rectangular shape into the slide
   */
  public int injectBasicShape(ShapeGeometry geometry, String text, String name) throws XMLParsingException {
    try {
      int spid = nextAvailableSpid++;

      // Create the shape element with all required child elements
      Element shapeElement = createBasicShapeElement(spid, name, geometry, text);

      // Inject into the shape tree
      shapeTree.appendChild(shapeElement);

      return spid;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to inject basic shape", e);
    }
  }

  /**
   * Update text content of an existing shape
   */
  public void updateShapeText(int spid, String newText) throws XMLParsingException {
    try {
      // Find the shape by SPID
      Element shape = findShapeBySpid(spid);
      if (shape == null) {
        throw new XMLParsingException("Shape with SPID " + spid + " not found");
      }

      // Find the text element and update it
      Element textElement = (Element) xpath.evaluate(".//a:t", shape, XPathConstants.NODE);
      if (textElement != null) {
        textElement.setTextContent(newText);
      } else {
        // Create text structure if it doesn't exist
        addTextToShape(shape, newText);
      }

    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to update shape text", e);
    }
  }

  /**
   * Update geometry (position/size) of an existing shape
   */
  public void updateShapeGeometry(int spid, ShapeGeometry newGeometry) throws XMLParsingException {
    try {
      Element shape = findShapeBySpid(spid);
      if (shape == null) {
        throw new XMLParsingException("Shape with SPID " + spid + " not found");
      }

      // Update the transform element
      updateShapeTransform(shape, newGeometry);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to update shape geometry", e);
    }
  }

  /**
   * Inject animation binding for a shape into the timing tree
   */
  public void injectAnimation(int targetSpid, String animationType, String transition, 
      String filter, String duration, String delay, int clickTrigger) throws XMLParsingException {
    try {
      // Find the timing tree
      Element timingElement = (Element) xpath.evaluate("//p:timing", document, XPathConstants.NODE);
      if (timingElement == null) {
        throw new XMLParsingException("No timing element found in slide");
      }

      // Find the specific click trigger node
      Element clickNode = findClickTriggerNode(timingElement, clickTrigger);
      if (clickNode == null) {
        // Create new click trigger if it doesn't exist
        clickNode = createNewClickTrigger(timingElement, clickTrigger);
      }

      // Create animation effect element
      Element animationEffect = createAnimationEffect(targetSpid, animationType, transition, filter, duration, delay);

      // Inject into the click trigger
      injectIntoClickTrigger(clickNode, animationEffect);

    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to inject animation", e);
    }
  }

  /**
   * Create a new click trigger in the timing sequence
   */
  public int createNewClickTrigger() throws XMLParsingException {
    try {
      Element timingElement = (Element) xpath.evaluate("//p:timing", document, XPathConstants.NODE);
      if (timingElement == null) {
        throw new XMLParsingException("No timing element found in slide");
      }

      // Find the main sequence
      Element mainSeq = (Element) xpath.evaluate(".//p:seq[@concurrent='1']//p:cTn", timingElement, XPathConstants.NODE);
      if (mainSeq == null) {
        throw new XMLParsingException("No main sequence found in timing");
      }

      // Get current click count and create new click trigger
      int newClickNumber = getNextClickTriggerNumber(mainSeq);
      Element newClickTrigger = createClickTriggerElement(newClickNumber);

      // Find or create childTnLst
      Element childTnLst = (Element) xpath.evaluate("./p:childTnLst", mainSeq, XPathConstants.NODE);
      if (childTnLst == null) {
        childTnLst = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:childTnLst");
        mainSeq.appendChild(childTnLst);
      }

      childTnLst.appendChild(newClickTrigger);
      return newClickNumber;

    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to create new click trigger", e);
    }
  }

  /**
   * Write the modified document to a file
   */
  public void writeXML(File outputFile) throws XMLParsingException {
    try {
      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();

      // Format the output nicely
      transformer.setOutputProperty(OutputKeys.INDENT, "yes");
      transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
      transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

      DOMSource source = new DOMSource(document);
      StreamResult result = new StreamResult(outputFile);

      transformer.transform(source, result);

    } catch (TransformerException e) {
      throw new XMLParsingException("Failed to write XML to file", e);
    }
  }

  /**
   * Create a complete basic shape element with all required OOXML structure
   */
  private Element createBasicShapeElement(int spid, String name, ShapeGeometry geometry, String text) {
    // Create the main shape element
    Element shape = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:sp");

    // Add non-visual properties
    Element nvSpPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvSpPr");
    shape.appendChild(nvSpPr);

    // Connection properties
    Element cNvPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvPr");
    cNvPr.setAttribute("id", String.valueOf(spid));
    cNvPr.setAttribute("name", name);
    nvSpPr.appendChild(cNvPr);

    // Shape connection properties
    Element cNvSpPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvSpPr");
    nvSpPr.appendChild(cNvSpPr);

    // Non-visual properties
    Element nvPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvPr");
    nvSpPr.appendChild(nvPr);

    // Shape properties
    Element spPr = createShapeProperties(geometry);
    shape.appendChild(spPr);

    // Add text body if text provided
    if (text != null && !text.trim().isEmpty()) {
      Element txBody = createTextBody(text);
      shape.appendChild(txBody);
    }

    return shape;
  }

  /**
   * Create shape properties element with geometry
   */
  private Element createShapeProperties(ShapeGeometry geometry) {
    Element spPr = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:spPr");

    // Transform element
    Element xfrm = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:xfrm");
    spPr.appendChild(xfrm);

    // Offset (position)
    Element off = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:off");
    off.setAttribute("x", String.valueOf(geometry.getX()));
    off.setAttribute("y", String.valueOf(geometry.getY()));
    xfrm.appendChild(off);

    // Extents (size)
    Element ext = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:ext");
    ext.setAttribute("cx", String.valueOf(geometry.getWidth()));
    ext.setAttribute("cy", String.valueOf(geometry.getHeight()));
    xfrm.appendChild(ext);

    // Preset geometry (rectangle)
    Element prstGeom = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:prstGeom");
    prstGeom.setAttribute("prst", "rect");
    spPr.appendChild(prstGeom);

    Element avLst = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:avLst");
    prstGeom.appendChild(avLst);

    return spPr;
  }

  /**
   * Create text body element with content
   */
  private Element createTextBody(String text) {
    Element txBody = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:txBody");

    // Body properties
    Element bodyPr = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:bodyPr");
    txBody.appendChild(bodyPr);

    // List style
    Element lstStyle = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:lstStyle");
    txBody.appendChild(lstStyle);

    // Paragraph
    Element p = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:p");
    txBody.appendChild(p);

    // Run
    Element r = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:r");
    p.appendChild(r);

    // Run properties
    Element rPr = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:rPr");
    rPr.setAttribute("lang", "en-US");
    rPr.setAttribute("sz", "1800"); // 18pt font
    r.appendChild(rPr);

    // Text
    Element t = document.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:t");
    t.setTextContent(text);
    r.appendChild(t);

    return txBody;
  }

  /**
   * Find a shape element by its SPID
   */
  private Element findShapeBySpid(int spid) throws XPathExpressionException {
    return (Element) xpath.evaluate("//p:sp[p:nvSpPr/p:cNvPr/@id='" + spid + "']", 
        document, XPathConstants.NODE);
  }

  /**
   * Calculate the next available SPID by finding the highest existing one
   */
  private int calculateNextSpid() throws XPathExpressionException {
    NodeList spids = (NodeList) xpath.evaluate("//p:cNvPr/@id", document, XPathConstants.NODESET);
    int maxSpid = 0;

    for (int i = 0; i < spids.getLength(); i++) {
      String spidStr = spids.item(i).getTextContent();
      try {
        int spid = Integer.parseInt(spidStr);
        maxSpid = Math.max(maxSpid, spid);
      } catch (NumberFormatException e) {
        // Skip non-numeric SPIDs
      }
    }

    return maxSpid + 1;
  }

  /**
   * Find a specific click trigger node in the timing tree
   */
  private Element findClickTriggerNode(Element timingElement, int clickNumber) throws XPathExpressionException {
    // Click triggers are par elements with indefinite delay
    NodeList clickTriggers = (NodeList) xpath.evaluate(
        ".//p:seq[@concurrent='1']//p:childTnLst/p:par[p:cTn/p:stCondLst/p:cond/@delay='indefinite']", 
        timingElement, XPathConstants.NODESET
        );

    if (clickNumber > 0 && clickNumber <= clickTriggers.getLength()) {
      return (Element) clickTriggers.item(clickNumber - 1);
    }

    return null;
  }

  /**
   * Create a new click trigger element
   */
  private Element createNewClickTrigger(Element timingElement, int clickNumber) throws XMLParsingException {
    try {
      // Find the main sequence
      Element mainSeq = (Element) xpath.evaluate(".//p:seq[@concurrent='1']//p:cTn", timingElement, XPathConstants.NODE);
      if (mainSeq == null) {
        throw new XMLParsingException("No main sequence found");
      }

      Element newClickTrigger = createClickTriggerElement(clickNumber);

      // Add to main sequence
      Element childTnLst = (Element) xpath.evaluate("./p:childTnLst", mainSeq, XPathConstants.NODE);
      if (childTnLst == null) {
        childTnLst = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:childTnLst");
        mainSeq.appendChild(childTnLst);
      }

      childTnLst.appendChild(newClickTrigger);
      return newClickTrigger;
    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to create new click trigger", e);
    }
  }

  /**
   * Create a click trigger element structure
   */
  private Element createClickTriggerElement(int clickNumber) {
    // Create par element for click trigger
    Element par = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:par");

    // Create cTn element
    Element cTn = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cTn");
    cTn.setAttribute("id", String.valueOf(getNextTimingNodeId()));
    cTn.setAttribute("fill", "hold");
    par.appendChild(cTn);

    // Create stCondLst for indefinite delay (click trigger)
    Element stCondLst = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:stCondLst");
    cTn.appendChild(stCondLst);

    Element cond = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cond");
    cond.setAttribute("delay", "indefinite");
    stCondLst.appendChild(cond);

    // Create childTnLst for animations
    Element childTnLst = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:childTnLst");
    cTn.appendChild(childTnLst);

    return par;
  }

  /**
   * Create an animation effect element
   */
  private Element createAnimationEffect(int targetSpid, String animationType, String transition, 
      String filter, String duration, String delay) {
    // Create par element for the animation
    Element par = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:par");

    // Create cTn element
    Element cTn = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cTn");
    cTn.setAttribute("id", String.valueOf(getNextTimingNodeId()));
    cTn.setAttribute("presetID", "10");
    cTn.setAttribute("presetClass", "in".equals(transition) ? "entr" : "exit");
    cTn.setAttribute("presetSubtype", "0");
    cTn.setAttribute("fill", "hold");
    cTn.setAttribute("nodeType", "clickEffect");
    par.appendChild(cTn);

    // Create stCondLst
    Element stCondLst = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:stCondLst");
    cTn.appendChild(stCondLst);

    Element cond = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cond");
    cond.setAttribute("delay", delay != null ? delay : "0");
    stCondLst.appendChild(cond);

    // Create childTnLst
    Element childTnLst = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:childTnLst");
    cTn.appendChild(childTnLst);

    // Create the animation effect
    Element animEffect = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:animEffect");
    animEffect.setAttribute("transition", transition);
    if (filter != null) {
      animEffect.setAttribute("filter", filter);
    }
    childTnLst.appendChild(animEffect);

    // Create cBhvr (common behavior)
    Element cBhvr = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cBhvr");
    animEffect.appendChild(cBhvr);

    // Create cTn for behavior
    Element behaviorCTn = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:cTn");
    behaviorCTn.setAttribute("id", String.valueOf(getNextTimingNodeId()));
    behaviorCTn.setAttribute("dur", duration != null ? duration : "330");
    cBhvr.appendChild(behaviorCTn);

    // Create target element
    Element tgtEl = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:tgtEl");
    cBhvr.appendChild(tgtEl);

    Element spTgt = document.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:spTgt");
    spTgt.setAttribute("spid", String.valueOf(targetSpid));
    tgtEl.appendChild(spTgt);

    return par;
  }

  /**
   * Update the transform element of a shape
   */
  private void updateShapeTransform(Element shape, ShapeGeometry geometry) throws XPathExpressionException {
    Element xfrm = (Element) xpath.evaluate(".//a:xfrm", shape, XPathConstants.NODE);
    if (xfrm == null) return;

    // Update offset
    Element off = (Element) xpath.evaluate("./a:off", xfrm, XPathConstants.NODE);
    if (off != null) {
      off.setAttribute("x", String.valueOf(geometry.getX()));
      off.setAttribute("y", String.valueOf(geometry.getY()));
    }

    // Update extents
    Element ext = (Element) xpath.evaluate("./a:ext", xfrm, XPathConstants.NODE);
    if (ext != null) {
      ext.setAttribute("cx", String.valueOf(geometry.getWidth()));
      ext.setAttribute("cy", String.valueOf(geometry.getHeight()));
    }
  }

  /**
   * Add text structure to a shape that doesn't have text
   */
  private void addTextToShape(Element shape, String text) {
    Element txBody = createTextBody(text);
    shape.appendChild(txBody);
  }

  /**
   * Inject animation into a click trigger
   */
  private void injectIntoClickTrigger(Element clickNode, Element animationEffect) throws XMLParsingException {
    try {
      Element childTnLst = (Element) xpath.evaluate("./p:cTn/p:childTnLst", clickNode, XPathConstants.NODE);
      if (childTnLst == null) {
        Element cTn = (Element) xpath.evaluate("./p:cTn", clickNode, XPathConstants.NODE);
        childTnLst = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:childTnLst");
        cTn.appendChild(childTnLst);
      }

      childTnLst.appendChild(animationEffect);
    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to inject animation into click trigger", e);
    }
  }

  /**
   * Get the next available timing node ID
   */
  private int getNextTimingNodeId() {
    try {
      NodeList timingIds = (NodeList) xpath.evaluate("//p:cTn/@id", document, XPathConstants.NODESET);
      int maxId = 0;

      for (int i = 0; i < timingIds.getLength(); i++) {
        String idStr = timingIds.item(i).getTextContent();
        try {
          int id = Integer.parseInt(idStr);
          maxId = Math.max(maxId, id);
        } catch (NumberFormatException e) {
          // Skip non-numeric IDs
        }
      }

      return maxId + 1;
    } catch (XPathExpressionException e) {
      // Return a safe default if XPath fails
      return 1000;
    }
  }

  /**
   * Get the next click trigger number
   */
  private int getNextClickTriggerNumber(Element mainSeq) throws XMLParsingException {
    try {
      NodeList clickTriggers = (NodeList) xpath.evaluate("./p:childTnLst/p:par", mainSeq, XPathConstants.NODESET);
      return clickTriggers.getLength() + 1;
    } catch (XPathExpressionException e) {
      throw new XMLParsingException("Failed to count existing click triggers", e);
    }
  }

  /**
   * Namespace context for XPath queries
   */
  private static class PowerPointNamespaceContext implements javax.xml.namespace.NamespaceContext {
    private static final Map<String, String> NAMESPACES = Map.of(
        "p", XMLConstants.PRESENTATION_NS,
        "a", XMLConstants.DRAWING_NS,
        "r", XMLConstants.RELATIONSHIPS_NS
        );

    @Override
    public String getNamespaceURI(String prefix) {
      return NAMESPACES.getOrDefault(prefix, javax.xml.XMLConstants.NULL_NS_URI);
    }

    @Override
    public String getPrefix(String namespaceURI) {
      return NAMESPACES.entrySet().stream()
        .filter(entry -> entry.getValue().equals(namespaceURI))
        .map(Map.Entry::getKey)
        .findFirst()
        .orElse(null);
    }

    @Override
    public java.util.Iterator<String> getPrefixes(String namespaceURI) {
      return NAMESPACES.entrySet().stream()
        .filter(entry -> entry.getValue().equals(namespaceURI))
        .map(Map.Entry::getKey)
        .iterator();
    }
  }
}
