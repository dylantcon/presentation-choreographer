package com.presentationchoreographer.core.model;

import org.w3c.dom.Element;

/**
 * Represents a shape or object on a PowerPoint slide
 */
public class SlideShape {
  public enum ShapeType {
    SHAPE, PICTURE, TEXT_BOX, GROUP
  }

  private final int spid;
  private final String name;
  private final ShapeType type;
  private final String textContent;
  private final ShapeGeometry geometry;
  private final Element xmlElement;

  public SlideShape(int spid, String name, ShapeType type, String textContent,
      ShapeGeometry geometry, Element xmlElement) {
    this.spid = spid;
    this.name = name;
    this.type = type;
    this.textContent = textContent;
    this.geometry = geometry;
    this.xmlElement = xmlElement;
  }

  // Getters
  public int getSpid() { return spid; }
  public String getName() { return name; }
  public ShapeType getType() { return type; }
  public String getTextContent() { return textContent; }
  public ShapeGeometry getGeometry() { return geometry; }
  public Element getXmlElement() { return xmlElement; }

  public boolean hasText() { return textContent != null && !textContent.trim().isEmpty(); }

  @Override
  public String toString() {
    return String.format("SlideShape{spid=%d, name='%s', type=%s, hasText=%s}",
        spid, name, type, hasText());
  }
}
