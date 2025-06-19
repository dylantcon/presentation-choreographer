package com.presentationchoreographer.core.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents a single node in the timing hierarchy
 */
public class TimingNode {
  private final String nodeId;
  private final String nodeType;
  private final String duration;
  private final List<TimingNode> children = new ArrayList<>();

  public TimingNode(String nodeId, String nodeType, String duration) {
    this.nodeId = nodeId;
    this.nodeType = nodeType;
    this.duration = duration;
  }

  public void addChild(TimingNode child) {
    children.add(child);
  }

  public String getNodeId() { return nodeId; }
  public String getNodeType() { return nodeType; }
  public String getDuration() { return duration; }
  public List<TimingNode> getChildren() { return new ArrayList<>(children); }

  public int getTotalNodeCount() {
    int count = 1; // This node
    for (TimingNode child : children) {
      count += child.getTotalNodeCount();
    }
    return count;
  }

  public void collectAllNodes(List<TimingNode> collector) {
    collector.add(this);
    for (TimingNode child : children) {
      child.collectAllNodes(collector);
    }
  }

  @Override
  public String toString() {
    return String.format("TimingNode{id='%s', type='%s', children=%d}",
        nodeId, nodeType, children.size());
  }
}
