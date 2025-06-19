package com.presentationchoreographer.core.model;

/**
 * Represents the position and size of a shape in EMUs (English Metric Units)
 */
public class ShapeGeometry {
  private final long x, y;      // Position (top-left corner)
  private final long width, height;  // Dimensions

  public ShapeGeometry(long x, long y, long width, long height) {
    this.x = x;
    this.y = y;
    this.width = width;
    this.height = height;
  }

  // Convert EMUs to points (common PowerPoint unit)
  public double getXInPoints() { return emuToPoints(x); }
  public double getYInPoints() { return emuToPoints(y); }
  public double getWidthInPoints() { return emuToPoints(width); }
  public double getHeightInPoints() { return emuToPoints(height); }

  // Raw EMU values
  public long getX() { return x; }
  public long getY() { return y; }
  public long getWidth() { return width; }
  public long getHeight() { return height; }

  // PowerPoint conversion: 1 point = 12700 EMUs
  private static double emuToPoints(long emu) {
    return emu / 12700.0;
  }

  @Override
  public String toString() {
    return String.format("Geometry{x=%.1fpt, y=%.1fpt, w=%.1fpt, h=%.1fpt}",
        getXInPoints(), getYInPoints(), getWidthInPoints(), getHeightInPoints());
  }
}
