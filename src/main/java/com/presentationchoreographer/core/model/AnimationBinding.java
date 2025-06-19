package com.presentationchoreographer.core.model;

/**
 * Represents the binding between an animation effect and its target shape
 */
public class AnimationBinding {
  private final int targetSpid;
  private final String animationType;
  private final String transition;
  private final String filter;
  private final String duration;
  private final String delay;

  public AnimationBinding(int targetSpid, String animationType, String transition,
      String filter, String duration, String delay) {
    this.targetSpid = targetSpid;
    this.animationType = animationType;
    this.transition = transition;
    this.filter = filter;
    this.duration = duration;
    this.delay = delay;
  }

  // Getters
  public int getTargetSpid() { return targetSpid; }
  public String getAnimationType() { return animationType; }
  public String getTransition() { return transition; }
  public String getFilter() { return filter; }
  public String getDuration() { return duration; }
  public String getDelay() { return delay; }

  public boolean isEntranceAnimation() {
    return "in".equals(transition);
  }

  public boolean isExitAnimation() {
    return "out".equals(transition);
  }

  @Override
  public String toString() {
    return String.format("AnimationBinding{spid=%d, type='%s', transition='%s', delay='%s'}",
        targetSpid, animationType, transition, delay);
  }
}
