# TODO - Presentation Choreographer

Development roadmap for completing the PowerPoint manipulation pipeline.

## Critical Path - PPTX Reconstruction Pipeline

### Priority 1: Relationship Management (.rels files)

**Status**: Implemented, Tested
**Complexity**: High  
**Impact**: Critical for valid PPTX files

#### Tasks:
- [x] **Implement `createSlideRelationships()`** in SlideCreator
  - Create `slide{N}.xml.rels` files for new slides
  - Standard relationships: slide layout, theme, master slide
  - Handle image/media relationships for slides with content
        
- [x] **Implement `copySlideRelationships()`** in SlideCreator
  - Copy source slide's `.rels` file to new location
  - Update relationship IDs to avoid conflicts
  - Remap media references if needed

- [x] **Relationship ID Management**
  - Create `RelationshipManager` class
  - Track and allocate unique relationship IDs across presentation
  - Update existing relationships when slides are inserted/moved

#### Technical Details:
```xml
<!-- Example slide1.xml.rels structure -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
```

---

### Priority 2: Presentation Structure Updates

**Status**: Implemented, unsure if tested
**Complexity**: Medium
**Impact**: Critical for slide ordering

#### Tasks:
- [x] **Implement `updatePresentationXml()`** in SlideCreator
  - Parse `ppt/presentation.xml`
  - Insert new `<p:sldId>` elements at correct position
  - Update `r:id` attributes sequentially
  - Maintain slide ordering integrity

- [x] **Slide ID Management**
  - Generate unique slide IDs for new slides
  - Update existing slide references in presentation.xml
  - Handle slide deletion scenarios (future)

#### Technical Details:
```xml
<!-- presentation.xml slide list section -->
<p:sldIdLst>
  <p:sldId id="256" r:id="rId2"/>
  <p:sldId id="257" r:id="rId3"/>  <!-- New slide inserted here -->
  <p:sldId id="258" r:id="rId4"/>  <!-- Subsequent slides renumbered -->
</p:sldIdLst>
```

---

### Priority 3: SPID Regeneration System

**Status**: Partially or wholly implemented, untested
**Complexity**: High  
**Impact**: Required for slide copying

#### Tasks:
- [x] **Implement `regenerateSpids()`** in SlideCreator
  - Parse all shape IDs in copied slide
  - Generate new unique SPIDs across entire presentation
  - Update shape references in animations/timing
  - Maintain internal slide consistency

- [x] **Global SPID Registry**
  - Create `SPIDManager` class
  - Track all SPIDs across all slides in presentation
  - Provide collision-free SPID allocation
  - Handle animation target updates

- [x] **Animation Reference Updates**
  - Update `<p:spTgt spid="...">` references when SPIDs change
  - Maintain timing tree integrity during SPID regeneration
  - Validate animation bindings after regeneration

#### Technical Complexity:
- Cross-slide SPID collision detection
- Animation timing preservation during ID changes
- Relationship maintenance between shapes and effects

---

### Priority 4: Content Types Registry

**Status**: Not Implemented  
**Complexity**: Low  
**Impact**: Required for OOXML compliance

#### Tasks:
- [ ] **Implement `updateContentTypes()`** in SlideCreator
  - Parse `[Content_Types].xml`
  - Ensure slide content types are registered
  - Add entries for new slides
  - Maintain OOXML compliance

#### Technical Details:
```xml
<!-- [Content_Types].xml structure -->
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
```

---

## Secondary Priorities - System Enhancement

### Priority 5: PPTX Orchestrator

**Status**: Design Phase  
**Complexity**: Medium  
**Impact**: End-to-end workflow completion

#### Tasks:
- [ ] **Create `PPTXOrchestrator` class**
  - Coordinate full extraction → modification → reconstruction pipeline
  - Handle ZIP compression/decompression
  - Validate OOXML structure integrity
  - Provide error recovery and rollback

- [ ] **ZIP Management**
  - Implement PowerShell/Java ZIP compression
  - Preserve file attributes and directory structure
  - Handle large presentations efficiently
  - Validate compressed output

- [ ] **End-to-End Testing**
  - Test complete workflow: .pptx → extract → modify → reconstruct → .pptx
  - Validate output in Microsoft PowerPoint
  - Performance testing with large presentations
  - Error handling and edge cases

---

### Priority 6: Template System Enhancement

**Status**: Foundation Complete  
**Complexity**: Medium  
**Impact**: User experience and automation

#### Tasks:
- [ ] **Implement `SlideTemplate` interface**
  - Create concrete template classes
  - Template data binding system
  - Reusable animation patterns
  - Content placeholder management

- [ ] **Concept Progression Template**
  - Based on your HTML → CSS → JavaScript pattern
  - Configurable timing intervals (330ms default)
  - Shape and animation coordination
  - Template data population

- [ ] **Phone Analogy Template**
  - Clock/Apps/Calls structure from your examples
  - Event-driven animation sequences
  - Customizable content and timing

---

### Priority 7: Animation Enhancement

**Status**: Core System Complete  
**Complexity**: Medium  
**Impact**: Advanced animation capabilities

#### Tasks:
- [ ] **Advanced Animation Effects**
  - Path-based animations (custom motion paths)
  - Transform animations (rotation, scaling)
  - Color and style animations
  - Complex effect combinations

- [ ] **Animation Templates**
  - Reusable animation sequences
  - Parameterized timing and effects
  - Animation library system
  - Effect chaining and coordination

- [ ] **Batch Animation Operations**
  - Apply animations to multiple shapes
  - Timing offset calculations
  - Mass animation updates
  - Animation pattern recognition

---

## Future Enhancements

### LLM Integration (Anthropic API)

**Status**: Planned  
**Complexity**: Medium  
**Impact**: Intelligent automation

#### Capabilities:
- **Content Generation**: Auto-generate slide content based on topics
- **Template Selection**: AI-powered template recommendations
- **Animation Intelligence**: Smart animation pattern selection
- **Content Analysis**: Understand slide relationships and flow

#### Implementation Areas:
- Template recommendation engine
- Content-aware animation selection  
- Automatic slide structuring
- Presentation flow optimization

---

### GUI Development

**Status**: Planned  
**Complexity**: High  
**Impact**: User accessibility

#### Framework Options:
- **JavaFX**: Native desktop application
- **Web Interface**: HTML/CSS/JavaScript with Java backend
- **Hybrid Approach**: Electron wrapper for cross-platform support

#### Core Features:
- Visual slide editor with drag-and-drop
- Animation timeline editor
- Template library browser
- Real-time PPTX preview
- Batch operation interface

---

### Performance Optimization

**Status**: Monitoring  
**Complexity**: Medium  
**Impact**: Scalability

#### Optimization Targets:
- **Memory Usage**: Streaming XML processing for large presentations
- **Processing Speed**: Parallel slide processing
- **File I/O**: Efficient ZIP operations
- **Cache Management**: Template and resource caching

---

## Implementation Strategy

### Phase 1: Complete Core Pipeline (Current Sprint)
1. Relationship management implementation
2. Presentation.xml updates
3. SPID regeneration system
4. Content types registry
5. End-to-end PPTX reconstruction testing

### Phase 2: Template System (Next Sprint)
1. Template interface implementation
2. Concrete template creation
3. Advanced animation capabilities
4. Template library foundation

### Phase 3: Intelligence Layer (Future Sprint)
1. LLM integration architecture
2. Content generation capabilities
3. Smart template recommendations
4. Automated presentation flow

### Phase 4: User Interface (Future Sprint)
1. GUI framework selection
2. Visual editor implementation
3. User experience optimization
4. Cross-platform deployment

---

## Technical Debt and Code Quality

### Current Issues:
- [ ] **Namespace Constants**: Complete XMLConstants integration in SlideXMLWriter
- [ ] **Exception Handling**: Standardize error messages and recovery paths
- [ ] **Code Documentation**: Add comprehensive JavaDoc comments
- [ ] **Unit Testing**: Create systematic test coverage
- [ ] **Configuration Management**: Externalize timing constants and settings

### Code Quality Improvements:
- [ ] **Design Patterns**: Apply Factory pattern for shape/animation creation
- [ ] **Dependency Injection**: Reduce tight coupling between components
- [ ] **Interface Segregation**: Split large interfaces into focused contracts
- [ ] **Performance Monitoring**: Add timing and memory usage tracking

---

## Success Criteria

### Milestone 1: Complete PPTX Pipeline
- [ ] Successfully create, modify, and reconstruct working .pptx files
- [ ] Validate output opens correctly in Microsoft PowerPoint
- [ ] All animations and content preserved through pipeline
- [ ] File size and structure comparable to PowerPoint-generated files

### Milestone 2: Template System
- [ ] Create slides from templates with data binding
- [ ] Apply complex animation patterns automatically
- [ ] Template library with 5+ production-ready templates
- [ ] Performance: Process 35+ slide presentations in < 10 seconds

### Milestone 3: Intelligence Integration
- [ ] LLM-powered content generation working
- [ ] Smart template selection based on content analysis
- [ ] Automated presentation flow optimization
- [ ] User satisfaction with AI-generated content quality

---

**Next Immediate Action**: Implement relationship management in SlideCreator to enable valid PPTX reconstruction.
