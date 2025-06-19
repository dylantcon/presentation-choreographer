package com.presentationchoreographer.utils;

/**
 * Constants for XML processing and PowerPoint namespaces
 */
public final class XMLConstants {

  // PowerPoint XML Namespaces
  public static final String PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
  public static final String DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  public static final String RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  // XPath expressions for shape extraction
  public static final String XPATH_ALL_SHAPES_AND_PICTURES = "//p:spTree//p:sp | //p:spTree//p:pic";
  public static final String XPATH_SHAPE_ID_ATTRIBUTE = ".//p:cNvPr/@id";
  public static final String XPATH_SHAPE_NAME_ATTRIBUTE = ".//p:cNvPr/@name";
  public static final String XPATH_SHAPE_TEXT_CONTENT = ".//a:t/text()";

  // XPath expressions for shape geometry
  public static final String XPATH_SHAPE_X_POSITION = ".//a:xfrm/a:off/@x";
  public static final String XPATH_SHAPE_Y_POSITION = ".//a:xfrm/a:off/@y";
  public static final String XPATH_SHAPE_WIDTH = ".//a:xfrm/a:ext/@cx";
  public static final String XPATH_SHAPE_HEIGHT = ".//a:xfrm/a:ext/@cy";

  // XPath expressions for timing and animation structure
  public static final String XPATH_TIMING_ROOT_ELEMENT = "//p:timing";
  public static final String XPATH_MAIN_ANIMATION_SEQUENCE = ".//p:seq[@concurrent='1']";
  public static final String XPATH_TIMING_CHILD_NODES = "./p:childTnLst/p:par | ./p:childTnLst/p:seq";
  public static final String XPATH_TIMING_CTN_ELEMENT = "./p:cTn";
  public static final String XPATH_TIMING_DELAY_ATTRIBUTE = "./p:stCondLst/p:cond/@delay";
  public static final String XPATH_TIMING_CTN_CHILDREN = "./p:childTnLst/p:par | ./p:childTnLst/p:seq";

  // XPath expressions for animation bindings
  public static final String XPATH_ALL_ANIMATION_EFFECTS = "//p:animEffect | //p:set";
  public static final String XPATH_ANIMATION_TARGET_SHAPE_ID = ".//p:spTgt/@spid";
  public static final String XPATH_ANIMATION_DURATION = "./p:cBhvr/p:cTn/@dur";
  public static final String XPATH_ANIMATION_DELAY = "./p:cBhvr/p:cTn/p:stCondLst/p:cond/@delay";

  // Animation timing constants
  public static final int DEFAULT_ANIMATION_INTERVAL_MS = 330;
  public static final String DEFAULT_ENTRANCE_EFFECT = "fade";
  public static final String DEFAULT_EXIT_EFFECT = "wipe(down)";

  // PowerPoint units conversion
  public static final double EMU_TO_POINTS = 12700.0;
  public static final double EMU_TO_INCHES = 914400.0;

  private XMLConstants() {
    // Utility class - no instantiation
  }
}
