# Presentation Choreographer

A comprehensive Java-based system for programmatic PowerPoint presentation creation and manipulation with surgical precision over OOXML structure, animations, and content.

## Project Overview

Presentation Choreographer provides developers and educators with fine-grained control over PowerPoint presentations through direct XML manipulation. The system enables automated slide creation, complex animation sequencing, template-based content generation, and intelligent presentation assembly.

### Key Capabilities

- **Surgical XML Control**: Direct manipulation of PowerPoint's OOXML structure
- **Shape Injection**: Programmatic creation and positioning of shapes with full geometry control
- **Animation Choreography**: Complex timing sequence creation with 330ms precision intervals
- **Slide Management**: Creation, duplication, and template-based slide generation
- **File System Pipeline**: Complete .pptx extraction, modification, and reconstruction workflow

## Architecture Overview

### Core Components

```
presentation-choreographer/
├── src/main/java/com/presentationchoreographer/
│   ├── core/
│   │   ├── model/           # Domain objects (Slide, Shape, Animation, Timing)
│   │   ├── timing/          # Animation timing engine
│   │   └── xml/             # Core XML processing abstractions
│   ├── xml/
│   │   ├── parsers/         # XML reading and parsing (SlideXMLParser)
│   │   ├── writers/         # XML generation and injection (SlideXMLWriter, SlideCreator)
│   │   └── validators/      # OOXML compliance validation
│   ├── templates/
│   │   ├── base/            # Abstract template classes
│   │   └── concrete/        # Specific templates (ConceptProgression, PhoneAnalogy)
│   ├── animations/
│   │   ├── sequences/       # Animation sequence definitions
│   │   ├── effects/         # Individual animation effects
│   │   └── bindings/        # Shape-to-animation relationships
│   ├── utils/               # XMLConstants and common utilities
│   └── exceptions/          # Custom exception hierarchy
└── test-pptx-samples/       # Test files and extracted OOXML content
```

### System Architecture

The system follows a layered architecture with clear separation of concerns:

1. **Parsing Layer** (`SlideXMLParser`): Extracts PowerPoint XML into structured Java objects
2. **Model Layer** (`core.model`): Represents slides, shapes, animations, and timing relationships
3. **Injection Layer** (`SlideXMLWriter`, `SlideCreator`): Programmatically generates and modifies XML
4. **Template Layer** (`templates`): Provides reusable slide and animation patterns
5. **Pipeline Layer**: Orchestrates the complete .pptx processing workflow

## Core Classes

### SlideXMLParser
Converts PowerPoint XML into navigable Java objects with comprehensive extraction of:
- **Shape Registry**: All shapes with positions, content, and metadata (21+ shapes parsed)
- **Timing Tree**: Hierarchical animation structure (64+ timing nodes, 4+ click triggers)
- **Animation Bindings**: Relationships between shapes and their animations (80+ bindings)

```java
SlideXMLParser parser = new SlideXMLParser();
ParsedSlideData slideData = parser.parseSlide(xmlFile);
```

### SlideXMLWriter
Provides surgical XML injection capabilities:
- **Shape Creation**: `injectBasicShape()` with precise geometry control
- **Animation Injection**: `injectAnimation()` with timing tree integration
- **Click Trigger Management**: `createNewClickTrigger()` for interactive sequences

```java
SlideXMLWriter writer = new SlideXMLWriter(document);
int spid = writer.injectBasicShape(geometry, "text", "shapeName");
writer.injectAnimation(spid, "p:animEffect", "in", "fade", "330", "0", clickTrigger);
```

### SlideCreator
Comprehensive slide management system supporting:
- **Blank Slide Creation**: `insertBlankSlide()` with minimal OOXML structure
- **Slide Duplication**: `insertCopiedSlide()` with conflict resolution
- **Template-Based Creation**: `insertTemplateSlide()` with data population
- **File System Cascade**: Automatic slide renaming and relationship management

```java
SlideCreator creator = new SlideCreator(extractedDir);
int newSlide = creator.insertBlankSlide(3, "New Slide Title");
```

## Data Model

### Core Entities

- **ParsedSlideData**: Container for all extracted slide information
- **SlideShape**: Individual shape with geometry, content, and metadata
- **ShapeGeometry**: Position and size information in EMUs (English Metric Units)
- **TimingTree**: Hierarchical structure of animation timing
- **TimingNode**: Individual timing elements with click triggers and delays
- **AnimationBinding**: Links between shapes and their animation effects

### Timing System

The system captures PowerPoint's sophisticated animation hierarchy:

```
MainSequence (Node 2)
├── Click Trigger 1 (Node 3) - indefinite delay
│   ├── Shape 19 fade in (330ms)
│   ├── Shape 25 fade in (660ms)
│   └── Shape 58 fade in (990ms)
├── Click Trigger 2 (Node 20) - indefinite delay
│   └── [6 sub-animations]
├── Click Trigger 3 (Node 45) - indefinite delay
│   └── [8 sub-animations]
└── Click Trigger 4 (Node 78) - indefinite delay
    └── [Exit sequence with wipe(down) effects]
    ```

## Current Test Results

### Shape Injection Verification
```
✅ SHAPE INJECTION SUCCESSFUL!
   → XML modification pipeline working correctly
      → Ready for animation injection development

      Shapes: 21 → 22 (1 added)
      Text shapes: 6 → 7 (1 added)
      New shape SPID: 66
      Position: Geometry{x=236.2pt, y=78.7pt, w=157.5pt, h=63.0pt}
      ```

### Animation Injection Verification
```
🎉 ANIMATION INJECTION SUCCESSFUL!
   → Shape injection working
      → Click trigger creation working
         → Animation binding injection working

         Animation bindings: 80 → 82 (2 added)
         Timing nodes: 64 → 68 (4 added)
         Click triggers: 4 → 6 (2 added)
         Animations for new shape: 2 (entrance + exit)
         ```

### Slide Creation Verification
```
✅ SLIDE CREATION SUCCESSFUL!
   → New slide file created: slide3.xml
      → File size: 751 bytes
         → Slide numbering cascade working

         Total slides: 35 → 36 (1 added)
         File system cascade: All subsequent slides renamed correctly
         ```

## Usage Examples

### Basic Shape and Animation Injection

```java
// Parse existing slide
SlideXMLParser parser = new SlideXMLParser();
ParsedSlideData slideData = parser.parseSlide(slideFile);

// Create XML writer
SlideXMLWriter writer = new SlideXMLWriter(document);

// Inject new shape
ShapeGeometry geometry = new ShapeGeometry(3000000L, 1000000L, 2000000L, 800000L);
int spid = writer.injectBasicShape(geometry, "My Shape", "Custom Shape");

// Add animations
int appearClick = writer.createNewClickTrigger();
writer.injectAnimation(spid, "p:animEffect", "in", "fade", "330", "0", appearClick);

int disappearClick = writer.createNewClickTrigger();
writer.injectAnimation(spid, "p:animEffect", "out", "wipe(down)", "500", "0", disappearClick);

// Write modified slide
writer.writeXML(outputFile);
```

### Slide Management

```java
// Create slide manager
SlideCreator creator = new SlideCreator(extractedPptxDir);

// Insert blank slide
int newSlide = creator.insertBlankSlide(3, "New Presentation Section");

// Copy existing slide
int copiedSlide = creator.insertCopiedSlide(5, 2, "Modified Copy of Slide 2");

// Create from template
TemplateData data = new TemplateData();
data.put("title", "Template-Generated Slide");
data.put("content", Arrays.asList("Point 1", "Point 2", "Point 3"));
int templateSlide = creator.insertTemplateSlide(7, conceptTemplate, data);
```

## Technical Specifications

### Dependencies
- **Java 22**: Modern Java features and performance
- **DOM XML Processing**: Native XML manipulation without external dependencies
- **XPath 1.0**: Query language for XML navigation
- **OOXML Standards**: Microsoft Office Open XML specification compliance

### Performance Characteristics
- **Memory Efficient**: DOM-based processing with selective loading
- **Precise Timing**: 330ms animation intervals with microsecond accuracy
- **Scalable Architecture**: Handles presentations with 35+ slides and 80+ animations
- **File System Optimized**: Efficient cascade operations for slide management

### Namespace Management
All XML operations use centralized namespace constants:
- `PRESENTATION_NS`: `http://schemas.openxmlformats.org/presentationml/2006/main`
- `DRAWING_NS`: `http://schemas.openxmlformats.org/drawingml/2006/main`
- `RELATIONSHIPS_NS`: `http://schemas.openxmlformats.org/officeDocument/2006/relationships`

## Testing Framework

The system includes comprehensive demos for each major component:

- **SlideXMLParserDemo**: Validates XML parsing and data extraction
- **ShapeInjectionDemo**: Tests shape creation and geometry management
- **AnimationInjectionDemo**: Verifies animation timing and click trigger creation
- **SlideCreationDemo**: Confirms slide management and file system operations

## Future Integration

The architecture is designed for future integration with:
- **Anthropic API**: LLM-powered content generation and template selection
- **GUI Framework**: JavaFX or web-based interface for visual editing
- **Template Marketplace**: Shareable animation and layout templates
- **Batch Processing**: Automated presentation generation pipelines

## Development Environment

### Requirements
- Java 22+ (tested with Java HotSpot 64-Bit Server VM)
- Windows 11 (primary development platform)
- PowerShell 7+ (for build scripts)
- Vim or IDE of choice (NetBeans recommended)

### Build Process
```powershell
# Compile core components
javac -d build src\main\java\com\presentationchoreographer\utils\*.java
javac -d build src\main\java\com\presentationchoreographer\exceptions\*.java
javac -d build src\main\java\com\presentationchoreographer\core\model\*.java
javac -d build src\main\java\com\presentationchoreographer\xml\parsers\*.java
javac -d build src\main\java\com\presentationchoreographer\xml\writers\*.java

# Run comprehensive tests
java -cp build com.presentationchoreographer.xml.AnimationInjectionDemo
java -cp build com.presentationchoreographer.xml.SlideCreationDemo
```

---

**Presentation Choreographer** - Bringing surgical precision to PowerPoint automation.
