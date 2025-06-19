package com.presentationchoreographer.xml.parsers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.xpath.*;
import javax.xml.namespace.NamespaceContext;
import javax.xml.XMLConstants;
import java.io.*;
import java.util.*;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.*;

/**
 * Core XML parser for PowerPoint slide files (.xml extracted from .pptx)
 * Handles DOM parsing and extraction of shapes, timing, and animation data
 */
public class SlideXMLParser {

  private final DocumentBuilder documentBuilder;
  private final XPath xpath;

  // Namespace mappings for PowerPoint XML
  private static final Map<String, String> NAMESPACES = Map.of(
      "p", com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS,
      "a", com.presentationchoreographer.utils.XMLConstants.DRAWING_NS,
      "r", com.presentationchoreographer.utils.XMLConstants.RELATIONSHIPS_NS
      );

  public SlideXMLParser() throws XMLParsingException {
    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();

      XPathFactory xpathFactory = XPathFactory.newInstance();
      this.xpath = xpathFactory.newXPath();
      this.xpath.setNamespaceContext(new PowerPointNamespaceContext());

    } catch (ParserConfigurationException e) {
      throw new XMLParsingException("Failed to initialize XML parser", e);
    }
  }

  /**
   * Parse a slide XML file and extract all critical data
   */
  public ParsedSlideData parseSlide(File xmlFile) throws XMLParsingException {
    try {
      Document document = documentBuilder.parse(xmlFile);
      return parseSlide(document);
    } catch (Exception e) {
      throw new XMLParsingException("Failed to parse slide XML file: " + xmlFile.getName(), e);
    }
  }

  /**
   * Parse a slide XML document and extract all critical data
   */
  public ParsedSlideData parseSlide(Document document) throws XMLParsingException {
    try {
      // Extract the three core data structures
      ShapeRegistry shapeRegistry = extractShapes(document);
      TimingTree timingTree = extractTimingTree(document);
      List<AnimationBinding> animationBindings = extractAnimationBindings(document, timingTree);

      return new ParsedSlideData(shapeRegistry, timingTree, animationBindings);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to parse slide document", e);
    }
  }

  /**
   * Extract all shapes from the slide with their spid mappings
   */
  private ShapeRegistry extractShapes(Document document) throws XPathExpressionException {
    ShapeRegistry registry = new ShapeRegistry();

    // Find all shapes in the shape tree
    NodeList shapeNodes = (NodeList) xpath.evaluate(
        com.presentationchoreographer.utils.XMLConstants.XPATH_ALL_SHAPES_AND_PICTURES, 
        document, 
        XPathConstants.NODESET
        );

    for (int i = 0; i < shapeNodes.getLength(); i++) {
      Element shapeElement = (Element) shapeNodes.item(i);
      SlideShape shape = parseShapeElement(shapeElement);
      if (shape != null) {
        registry.addShape(shape);
      }
    }

    return registry;
  }

  /**
   * Parse an individual shape element
   */
  private SlideShape parseShapeElement(Element shapeElement) throws XPathExpressionException {
    // Extract spid from cNvPr element
    String spidStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_ID_ATTRIBUTE, shapeElement, XPathConstants.STRING);
    if (spidStr.isEmpty()) return null;

    int spid = Integer.parseInt(spidStr);

    // Extract shape name
    String name = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_NAME_ATTRIBUTE, shapeElement, XPathConstants.STRING);

    // Determine shape type
    String tagName = shapeElement.getTagName();
    SlideShape.ShapeType type = tagName.equals("p:pic") ? 
      SlideShape.ShapeType.PICTURE : SlideShape.ShapeType.SHAPE;

    // Extract text content if present
    String textContent = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_TEXT_CONTENT, shapeElement, XPathConstants.STRING);

    // Extract position and size
    ShapeGeometry geometry = extractShapeGeometry(shapeElement);

    return new SlideShape(spid, name, type, textContent, geometry, shapeElement);
  }

  /**
   * Extract shape position and size information
   */
  private ShapeGeometry extractShapeGeometry(Element shapeElement) throws XPathExpressionException {
    String xStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_X_POSITION, shapeElement, XPathConstants.STRING);
    String yStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_Y_POSITION, shapeElement, XPathConstants.STRING);
    String cxStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_WIDTH, shapeElement, XPathConstants.STRING);
    String cyStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_SHAPE_HEIGHT, shapeElement, XPathConstants.STRING);

    // PowerPoint uses EMUs (English Metric Units)
    long x = xStr.isEmpty() ? 0 : Long.parseLong(xStr);
    long y = yStr.isEmpty() ? 0 : Long.parseLong(yStr);
    long cx = cxStr.isEmpty() ? 0 : Long.parseLong(cxStr);
    long cy = cyStr.isEmpty() ? 0 : Long.parseLong(cyStr);

    return new ShapeGeometry(x, y, cx, cy);
  }

  /**
   * Extract the complete timing tree structure
   */
  private TimingTree extractTimingTree(Document document) throws XPathExpressionException {
    Element timingElement = (Element) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_TIMING_ROOT_ELEMENT, document, XPathConstants.NODE);
    if (timingElement == null) {
      return new TimingTree(); // Empty timing tree
    }

    return parseTimingElement(timingElement);
  }

  /**
   * Parse the timing element into a structured tree
   */
  private TimingTree parseTimingElement(Element timingElement) throws XPathExpressionException {
    TimingTree tree = new TimingTree();

    // Find the main sequence
    Element mainSeq = (Element) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_MAIN_ANIMATION_SEQUENCE, timingElement, XPathConstants.NODE);
    if (mainSeq != null) {
      TimingNode rootNode = parseTimingNode(mainSeq);
      tree.setRootNode(rootNode);
    }

    return tree;
  }

  /**
   * Recursively parse timing nodes
   */
  private TimingNode parseTimingNode(Element element) throws XPathExpressionException {
    String nodeId = element.getAttribute("id");
    String nodeType = element.getAttribute("nodeType");
    String duration = element.getAttribute("dur");

    TimingNode node = new TimingNode(nodeId, nodeType, duration);

    // Parse child timing nodes
    NodeList children = (NodeList) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_TIMING_CHILD_NODES, 
        element, XPathConstants.NODESET);

    for (int i = 0; i < children.getLength(); i++) {
      Element childElement = (Element) children.item(i);
      TimingNode childNode = parseTimingNode(childElement);
      node.addChild(childNode);
    }

    return node;
  }

  /**
   * Extract animation bindings (which animations target which shapes)
   */
  private List<AnimationBinding> extractAnimationBindings(Document document, TimingTree timingTree) 
      throws XPathExpressionException {

      List<AnimationBinding> bindings = new ArrayList<>();

      // Find all animation effects that target shapes
      NodeList animEffects = (NodeList) xpath.evaluate(
          com.presentationchoreographer.utils.XMLConstants.XPATH_ALL_ANIMATION_EFFECTS, 
          document, 
          XPathConstants.NODESET
          );

      for (int i = 0; i < animEffects.getLength(); i++) {
        Element effectElement = (Element) animEffects.item(i);
        AnimationBinding binding = parseAnimationBinding(effectElement);
        if (binding != null) {
          bindings.add(binding);
        }
      }

      return bindings;
  }

  /**
   * Parse an individual animation binding
   */
  private AnimationBinding parseAnimationBinding(Element effectElement) throws XPathExpressionException {
    // Extract target shape ID
    String spidStr = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_ANIMATION_TARGET_SHAPE_ID, effectElement, XPathConstants.STRING);
    if (spidStr.isEmpty()) return null;

    int targetSpid = Integer.parseInt(spidStr);

    // Extract animation type and properties
    String animationType = effectElement.getTagName();
    String transition = effectElement.getAttribute("transition");
    String filter = effectElement.getAttribute("filter");

    // Extract timing information
    String duration = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_ANIMATION_DURATION, effectElement, XPathConstants.STRING);
    String delay = (String) xpath.evaluate(com.presentationchoreographer.utils.XMLConstants.XPATH_ANIMATION_DELAY, 
        effectElement, XPathConstants.STRING);

    return new AnimationBinding(targetSpid, animationType, transition, filter, duration, delay);
  }

  /**
   * Namespace context for XPath queries
   */
  private static class PowerPointNamespaceContext implements NamespaceContext {
    @Override
    public String getNamespaceURI(String prefix) {
      return NAMESPACES.getOrDefault(prefix, XMLConstants.NULL_NS_URI);
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
    public Iterator<String> getPrefixes(String namespaceURI) {
      return NAMESPACES.entrySet().stream()
        .filter(entry -> entry.getValue().equals(namespaceURI))
        .map(Map.Entry::getKey)
        .iterator();
    }
  }
}
