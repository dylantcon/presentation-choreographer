package com.presentationchoreographer.core.model;

import java.util.List;

/**
 * Container for all parsed data from a slide XML
 */
public class ParsedSlideData {
  private final ShapeRegistry shapeRegistry;
  private final TimingTree timingTree;
  private final List<AnimationBinding> animationBindings;

  public ParsedSlideData(ShapeRegistry shapeRegistry, TimingTree timingTree, 
      List<AnimationBinding> animationBindings) {
    this.shapeRegistry = shapeRegistry;
    this.timingTree = timingTree;
    this.animationBindings = animationBindings;
  }

  public ShapeRegistry getShapeRegistry() { return shapeRegistry; }
  public TimingTree getTimingTree() { return timingTree; }
  public List<AnimationBinding> getAnimationBindings() { return animationBindings; }

  @Override
  public String toString() {
    return String.format("ParsedSlideData{shapes=%d, timingNodes=%d, animations=%d}",
        shapeRegistry.getShapeCount(),
        timingTree.getNodeCount(),
        animationBindings.size());
  }
}
