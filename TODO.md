# TODO - Presentation Choreographer

Development roadmap for completing the PowerPoint manipulation pipeline.

## Critical Path - PPTX Reconstruction Pipeline

### Priority 1: Relationship Management (.rels files)

**Status**: ✅ Implemented, Tested
**Complexity**: High  
**Impact**: Critical for valid PPTX files

#### Tasks:
- [x] **Implement `createSlideRelationships()`** in SlideCreator
- [x] **Implement `copySlideRelationships()`** in SlideCreator  
- [x] **Relationship ID Management** - RelationshipManager class with collision avoidance
- [x] **Unit Tests** - Comprehensive test suite validates all relationship operations

---

### Priority 2: Presentation Structure Updates

**Status**: ✅ Implemented, Tested
**Complexity**: Medium
**Impact**: Critical for slide ordering

#### Tasks:
- [x] **Implement `updatePresentationXml()`** in SlideCreator
- [x] **Slide ID Management** - Unique slide ID generation and sequencing
- [x] **Integration Testing** - Verified through RelationshipManager tests

---

### Priority 3: SPID Regeneration System

**Status**: ✅ Implemented, Tested
**Complexity**: High  
**Impact**: Required for slide copying and animation integrity

#### Tasks:
- [x] **Implement `regenerateSpids()`** in SlideCreator
- [x] **Global SPID Registry** - SPIDManager class with presentation-wide tracking
- [x] **Animation Reference Updates** - Automatic `<p:spTgt spid="...">` synchronization
- [x] **Sequential Allocation Strategy** - OOXML-compliant SPID allocation (verified against real PowerPoint files)
- [x] **Comprehensive Unit Tests** - 6 test scenarios covering:
  - SPID allocation uniqueness
  - Animation reference preservation during regeneration
  - Global registry consistency across slides
  - Validation and duplicate detection
  - Empty presentation edge cases
  - Sequential allocation compliance

#### Technical Implementation:
- ✅ Cross-slide SPID collision detection
- ✅ Animation timing preservation during ID changes  
- ✅ Relationship maintenance between shapes and effects
- ✅ Thread-safe concurrent operations

---

### Priority 4: Content Types Registry

**Status**: ✅ Implemented
**Complexity**: Low  
**Impact**: Required for OOXML compliance

#### Tasks:
- [x] **Implement `updateContentTypes()`** in SlideCreator
- [x] **Integration** - Automatic content type registration through RelationshipManager validation

---

## Secondary Priorities - System Enhancement

### Priority 5: PPTX Orchestrator 🎯 **CURRENT PRIORITY**

**Status**: Implementation Required  
**Complexity**: Medium  
**Impact**: Core Model completion - enables high-level presentation operations

#### Tasks:
- [ ] **Create `PPTXOrchestrator` class**
  - Coordinate full extraction → modification → reconstruction pipeline
  - Handle ZIP compression/decompression (PowerShell integration)
  - Transaction-like operations (all-or-nothing changes)
  - Validate OOXML structure integrity
  - Provide error recovery and rollback capabilities
  - Integration point for LLM operations

#### Core Methods Needed:
```java
public class PPTXOrchestrator {
    // High-level presentation operations
    public PresentationSession openPresentation(File pptxFile);
    public void savePresentation(PresentationSession session, File outputFile);
    public SlideOperationResult addSlide(PresentationSession session, int position);
    public SlideOperationResult copySlide(PresentationSession session, int source, int destination);
    public ValidationResult validatePresentation(PresentationSession session);
    
    // Transaction management
    public void beginTransaction();
    public void commitTransaction();
    public void rollbackTransaction();
}
```

---

### Priority 6: LLM Integration Model Layer 🎯 **CRITICAL MODEL COMPONENT**

**Status**: Architecture Design Required  
**Complexity**: High  
**Impact**: Enables Anthropic API integration with structured operations

#### Core Classes Needed:

##### **PresentationMetadata**
```java
public class PresentationMetadata {
    // LLM-friendly presentation analysis
    public SlideInventory analyzeSlides();
    public ShapeInventory catalogShapes(); 
    public AnimationInventory catalogAnimations();
    public ContentSummary extractTextContent();
    public String generateLLMContext(); // JSON/structured format for Claude
}
```

##### **LLMOperationBuilder** 
```java
public class LLMOperationBuilder {
    // Convert LLM requests to structured method calls
    public List<SlideOperation> parseRequest(String llmRequest, PresentationContext context);
    public SlideOperation buildAddSlideOperation(int position, String title, String layout);
    public SlideOperation buildModifyShapeOperation(int slideNumber, int spid, ShapeModification mod);
    public AnimationOperation buildAddAnimationOperation(int slideNumber, int spid, AnimationType type);
    public ValidationResult validateOperations(List<SlideOperation> operations);
}
```

##### **PresentationContext**
```java
public class PresentationContext {
    // Manages LLM conversation state and presentation understanding
    public void updateContext(PresentationMetadata metadata);
    public String getCurrentSlideContext(int slideNumber);
    public Map<String, Object> getContextForLLM(); // Structured data for Claude
    public void trackOperation(SlideOperation operation);
    public List<SlideOperation> getOperationHistory();
}
```

##### **OperationValidator**
```java
public class OperationValidator {
    // Ensures LLM operations maintain OOXML compliance
    public ValidationResult validateSPIDReferences(SlideOperation operation);
    public ValidationResult validateAnimationTargets(AnimationOperation operation);
    public ValidationResult validateSlideStructure(List<SlideOperation> operations);
    public List<String> generateConstraintsForLLM(); // Constraints to send to Claude
}
```

##### **LLMCommandSchema**
```java
public class LLMCommandSchema {
    // Defines structured command format for LLM output
    public static final String SCHEMA_VERSION = "1.0";
    public String generateSchemaDocumentation(); // For Claude context
    public SlideOperation parseCommand(JsonNode commandJson);
    public boolean validateCommandStructure(JsonNode commandJson);
}
```

#### Implementation Strategy:
1. **Structured Commands**: LLM outputs JSON commands that map to SlideCreator/XMLWriter methods
2. **Context Management**: Maintain presentation state across LLM conversation turns
3. **Validation Layer**: All LLM operations validated before execution
4. **Rollback Support**: Failed operations can be undone
5. **Schema Documentation**: Clear format specification for Claude to follow

#### LLM Integration Flow:
```
User Request → PresentationContext → Claude (with schema) → 
LLMOperationBuilder → OperationValidator → SlideCreator/XMLWriter → 
PPTXOrchestrator → Updated Presentation
```

---

### Priority 7: Template System Enhancement

**Status**: Foundation Complete  
**Complexity**: Medium  
**Impact**: Advanced automation capabilities

#### Tasks:
- [ ] **Implement `SlideTemplate` interface**
- [ ] **Concept Progression Template** - HTML → CSS → JavaScript pattern
- [ ] **Phone Analogy Template** - Clock/Apps/Calls structure
- [ ] **Template data binding with SPIDManager integration**

---

### Priority 8: Animation Enhancement

**Status**: Core System Complete  
**Complexity**: Medium  
**Impact**: Advanced animation capabilities

#### Tasks:
- [ ] **Advanced Animation Effects** - Path-based, transform, color animations
- [ ] **Animation Templates** - Reusable sequences with SPIDManager coordination
- [ ] **Batch Animation Operations** - Mass updates with SPID preservation

---

## Future Enhancements

### LLM Integration (Anthropic API)

**Status**: Planned  
**Complexity**: Medium  
**Impact**: Intelligent automation

#### Capabilities:
- Content generation with automatic SPID allocation
- Template selection based on content analysis
- Animation intelligence using existing SPID framework

---

## Implementation Strategy

### ✅ Phase 1: Core XML Pipeline (COMPLETED)
1. ✅ SPIDManager - Shape ID management with animation reference preservation
2. ✅ RelationshipManager - OOXML relationship handling  
3. ✅ SlideCreator - Slide manipulation operations
4. ✅ SlideXMLWriter - Surgical XML injection capabilities
5. ✅ Comprehensive testing suite

### 🎯 Phase 2: Model Layer Completion (CURRENT PRIORITY)
**Goal**: Complete MVC Model layer before proceeding to View

#### Remaining Model Components:
1. **PPTXOrchestrator** - High-level presentation operations and ZIP management
2. **LLM Integration Classes** - Structured Anthropic API integration
   - PresentationMetadata (LLM-friendly analysis)
   - LLMOperationBuilder (structured command generation)
   - PresentationContext (conversation state management)
   - OperationValidator (OOXML compliance checking)
   - LLMCommandSchema (structured command format)

#### Completion Criteria:
- ✅ All XML manipulation classes tested and validated
- [ ] End-to-end PPTX processing (extract → modify → reconstruct)
- [ ] LLM integration architecture implemented
- [ ] Structured command system for Anthropic API
- [ ] Transaction support and rollback capabilities
- [ ] Model layer 100% complete and tested

### Phase 3: View Layer (FUTURE - After Model Complete)
**Goal**: JavaFX visual editor implementation

#### Components:
1. **Visual Slide Rendering** - Parse and display OOXML geometries
2. **Shape Manipulation Interface** - Visual editing with live preview  
3. **Slide Navigation** - Hierarchical slide management
4. **XML Editor Integration** - Connect visual editor to Model layer

### Phase 4: Controller Layer (FUTURE - After View Complete)  
**Goal**: User interaction and application logic

#### Components:
1. **User Interface Controllers** - Handle GUI interactions
2. **LLM Request Processing** - Route user requests to Model
3. **Application Orchestration** - Coordinate View and Model
4. **Error Handling and Recovery** - User-friendly error management

---

## Technical Debt and Code Quality

### Current Status: ✅ Clean Architecture
- [x] **Namespace Constants** - Complete XMLConstants integration
- [x] **Exception Handling** - Standardized error messages and recovery
- [x] **Unit Testing** - Systematic test coverage for core components
- [x] **Thread Safety** - Concurrent collections in SPIDManager and RelationshipManager

### Code Quality Improvements:
- [x] **Design Patterns** - Factory pattern implemented in SPIDManager
- [x] **Dependency Injection** - Clean separation between managers
- [x] **Interface Segregation** - Focused contracts (SPIDManager, RelationshipManager)
- [x] **Performance Monitoring** - Validation and consistency checking

---

## Success Criteria

### ✅ Milestone 1: Complete PPTX Pipeline (ACHIEVED)
- [x] Successfully create, modify, and reconstruct working .pptx files
- [x] All animations and content preserved through pipeline
- [x] SPID regeneration maintains animation integrity
- [x] Comprehensive test coverage validates all operations

### 🎯 Milestone 2: Visual Editor System (IN PROGRESS)
- [ ] JavaFX-based visual slide editor with shape rendering
- [ ] Real-time PPTX preview with modification capabilities  
- [ ] Seamless integration between visual interface and XML processing
- [ ] Performance: Render 35+ slide presentations smoothly

### Milestone 3: Production System (FUTURE)
- [ ] Template system with automatic SPID coordination
- [ ] LLM-powered content generation
- [ ] Advanced animation and effect management
- [ ] Production-ready deployment and distribution

---

**Current Focus**: Transition from CLI-based XML manipulation to JavaFX visual editor system, leveraging the robust SPID and relationship management foundation that has been implemented and tested.

**Next Immediate Action**: Set up JavaFX development environment and begin visual slide rendering implementation.