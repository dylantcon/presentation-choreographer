package com.presentationchoreographer.xml;

import com.presentationchoreographer.xml.parsers.SlideXMLParser;
import com.presentationchoreographer.xml.writers.SlideXMLWriter;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.XMLParsingException;
import javax.xml.parsers.*;
import org.w3c.dom.*;
import java.io.File;

/**
 * Comprehensive test of animation injection capabilities
 * Tests the XML Writer's ability to inject animations into timing hierarchy
 */
public class AnimationInjectionDemo {

  public static void main(String[] args) {
    try {
      System.out.println("=== Animation Injection Demo ===");
      System.out.println("Testing animation modification capabilities...");
      System.out.println();

      // Load the original slide
      File originalSlide = new File("test-pptx-samples/extracted/ppt/slides/slide2.xml");
      if (!originalSlide.exists()) {
        System.err.println("Original slide not found!");
        return;
      }

      // Parse the original slide
      SlideXMLParser parser = new SlideXMLParser();
      ParsedSlideData originalData = parser.parseSlide(originalSlide);

      System.out.println("BEFORE INJECTION:");
      System.out.println("  Shapes: " + originalData.getShapeRegistry().getShapeCount());
      System.out.println("  Animation bindings: " + originalData.getAnimationBindings().size());
      System.out.println("  Timing nodes: " + originalData.getTimingTree().getNodeCount());
      System.out.println("  Click triggers: " + countClickTriggers(originalData.getTimingTree()));
      System.out.println();

      // Load document for modification
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      DocumentBuilder builder = factory.newDocumentBuilder();
      Document document = builder.parse(originalSlide);

      // Create the XML writer
      SlideXMLWriter writer = new SlideXMLWriter(document);

      // Step 1: Inject a new test shape
      System.out.println("STEP 1: INJECTING TEST SHAPE...");
      ShapeGeometry testGeometry = new ShapeGeometry(
          2500000L,  // X position
          6000000L,  // Y position (bottom of slide)
          1500000L,  // Width
          600000L    // Height
          );

      int newSpid = writer.injectBasicShape(
          testGeometry, 
          "ANIMATED SQUARE", 
          "Test Animation Target"
          );
      System.out.println("  ✓ New shape created with SPID: " + newSpid);
      System.out.println();

      // Step 2: Create new click triggers for "Appear" and "Disappear"
      System.out.println("STEP 2: CREATING NEW CLICK TRIGGERS...");

      // Create "Appear" click trigger (entrance animation)
      int appearClickTrigger = writer.createNewClickTrigger();
      writer.injectAnimation(newSpid, "p:animEffect", "in", "fade", "330", "0", appearClickTrigger);
      System.out.println("  ✓ Created 'Appear' click trigger: " + appearClickTrigger);

      // Create "Disappear" click trigger (exit animation)  
      int disappearClickTrigger = writer.createNewClickTrigger();
      writer.injectAnimation(newSpid, "p:animEffect", "out", "wipe(down)", "500", "0", disappearClickTrigger);
      System.out.println("  ✓ Created 'Disappear' click trigger: " + disappearClickTrigger);
      System.out.println();

      // Step 3: Write the modified slide
      File modifiedSlide = new File("test-pptx-samples/slide2_with_animations.xml");
      writer.writeXML(modifiedSlide);

      System.out.println("ANIMATION INJECTION COMPLETE:");
      System.out.println("  Modified slide saved to: " + modifiedSlide.getName());
      System.out.println();

      // Step 4: Verify changes
      ParsedSlideData modifiedData = parser.parseSlide(modifiedSlide);

      System.out.println("AFTER INJECTION:");
      System.out.println("  Shapes: " + modifiedData.getShapeRegistry().getShapeCount());
      System.out.println("  Animation bindings: " + modifiedData.getAnimationBindings().size());
      System.out.println("  Timing nodes: " + modifiedData.getTimingTree().getNodeCount());
      System.out.println("  Click triggers: " + countClickTriggers(modifiedData.getTimingTree()));
      System.out.println();

      // Verify our injected shape exists and has animations
      SlideShape injectedShape = modifiedData.getShapeRegistry().getShape(newSpid);
      if (injectedShape != null) {
        System.out.println("INJECTED SHAPE VERIFICATION:");
        System.out.println("  ✓ Shape found with SPID: " + injectedShape.getSpid());
        System.out.println("  ✓ Name: " + injectedShape.getName());
        System.out.println("  ✓ Text: \"" + injectedShape.getTextContent() + "\"");
        System.out.println();
      }

      // Count animations targeting our new shape
      long animationsForNewShape = modifiedData.getAnimationBindings().stream()
        .filter(binding -> binding.getTargetSpid() == newSpid)
        .count();

      System.out.println("ANIMATION VERIFICATION:");
      System.out.println("  Animations targeting new shape: " + animationsForNewShape);

      // Show details of animations for our shape
      modifiedData.getAnimationBindings().stream()
        .filter(binding -> binding.getTargetSpid() == newSpid)
        .forEach(binding -> {
          System.out.printf("    → %s (%s) - delay: %s, duration: %s%n",
              binding.getTransition(),
              binding.getFilter(),
              binding.getDelay(),
              binding.getDuration());
        });

      System.out.println();

      // Calculate changes
      int shapeIncrease = modifiedData.getShapeRegistry().getShapeCount() - 
        originalData.getShapeRegistry().getShapeCount();
      int animationIncrease = modifiedData.getAnimationBindings().size() - 
        originalData.getAnimationBindings().size();
      int timingNodeIncrease = modifiedData.getTimingTree().getNodeCount() - 
        originalData.getTimingTree().getNodeCount();
      int clickTriggerIncrease = countClickTriggers(modifiedData.getTimingTree()) - 
        countClickTriggers(originalData.getTimingTree());

      System.out.println("=== VERIFICATION RESULTS ===");
      System.out.printf("Shape increase: %d (expected: 1)%n", shapeIncrease);
      System.out.printf("Animation increase: %d (expected: 4+) [2 effects × 2 bindings each]%n", animationIncrease);
      System.out.printf("Timing node increase: %d (expected: 10+)%n", timingNodeIncrease);
      System.out.printf("Click trigger increase: %d (expected: 2)%n", clickTriggerIncrease);
      System.out.printf("Animations for new shape: %d (expected: 2+)%n", animationsForNewShape);

      System.out.println();

      if (shapeIncrease == 1 && clickTriggerIncrease == 2 && animationsForNewShape >= 2) {
        System.out.println("🎉 ANIMATION INJECTION SUCCESSFUL!");
        System.out.println("   → Shape injection working");
        System.out.println("   → Click trigger creation working");
        System.out.println("   → Animation binding injection working");
        System.out.println("   → Ready for slide creation and PPTX reconstruction");
      } else {
        System.out.println("❌ ANIMATION INJECTION ISSUES DETECTED");
        System.out.println("   → Check the results above for specific issues");
      }

    } catch (XMLParsingException e) {
      System.err.println("XML Parsing Error: " + e.getMessage());
      e.printStackTrace();
    } catch (Exception e) {
      System.err.println("Unexpected Error: " + e.getMessage());
      e.printStackTrace();
    }
  }

  /**
   * Count click triggers in the timing tree (nodes with delay="indefinite")
   */
  private static int countClickTriggers(TimingTree timingTree) {
    return (int) timingTree.getAllNodes().stream()
      .filter(TimingNode::isClickTrigger)
      .count();
  }
}
