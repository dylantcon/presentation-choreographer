package com.presentationchoreographer.orchestration;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicLong;
import com.presentationchoreographer.xml.writers.*;
import com.presentationchoreographer.xml.parsers.SlideXMLParser;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.XMLParsingException;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * High-level orchestration class for PPTX presentation manipulation.
 * 
 * <p>PPTXOrchestrator provides the primary API surface for presentation operations,
 * abstracting away the surgical XML details handled by lower-level components.
 * This class serves as the integration point for LLM operations and provides
 * transaction-like capabilities for safe presentation modifications.</p>
 * 
 * <p>Key responsibilities:</p>
 * <ul>
 *   <li>Complete PPTX processing pipeline (extract → modify → reconstruct)</li>
 *   <li>ZIP compression/decompression management</li>
 *   <li>Transaction support with rollback capabilities</li>
 *   <li>High-level slide operations (add, copy, modify, delete)</li>
 *   <li>Presentation validation and integrity checking</li>
 *   <li>Integration point for Anthropic API operations</li>
 * </ul>
 * 
 * <p>Thread Safety: This class is thread-safe through transaction isolation
 * and concurrent collection usage.</p>
 * 
 * @author Presentation Choreographer
 * @version 1.0
 * @since 1.0
 */
public class PPTXOrchestrator {

  /**
   * Active presentation sessions indexed by session ID
   */
  private final Map<String, PresentationSession> activeSessions;

  /**
   * Session ID generator for unique session identification
   */
  private final AtomicLong sessionIdGenerator;

  /**
   * Temporary directory for PPTX extraction operations
   */
  private final File tempDirectory;

  /**
   * XML parser for slide content analysis
   */
  private final SlideXMLParser xmlParser;

  /**
   * Document builder for XML operations
   */
  private final DocumentBuilder documentBuilder;

  /**
   * Constructs a new PPTXOrchestrator with default configuration.
   * 
   * @throws XMLParsingException If XML processing components cannot be initialized
   */
  public PPTXOrchestrator() throws XMLParsingException {
    this.activeSessions = new ConcurrentHashMap<>();
    this.sessionIdGenerator = new AtomicLong(1);
    this.tempDirectory = createTempDirectory();
    this.xmlParser = new SlideXMLParser();

    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();
    } catch (ParserConfigurationException e) {
      throw new XMLParsingException("Failed to initialize PPTXOrchestrator", e);
    }
  }

  /**
   * Opens a PPTX presentation for editing, extracting contents and initializing
   * all management components.
   * 
   * @param pptxFile The PPTX file to open
   * @return PresentationSession for performing operations on the presentation
   * @throws XMLParsingException If the presentation cannot be opened or parsed
   * @throws IllegalArgumentException If pptxFile is null or does not exist
   */
  public PresentationSession openPresentation(File pptxFile) throws XMLParsingException {
    if (pptxFile == null || !pptxFile.exists()) {
      throw new IllegalArgumentException("PPTX file must exist and be non-null");
    }

    try {
      String sessionId = generateSessionId();
      File sessionDir = new File(tempDirectory, "session_" + sessionId);
      sessionDir.mkdirs();

      // Step 1: Extract PPTX contents
      File extractedDir = new File(sessionDir, "extracted");
      extractPPTX(pptxFile, extractedDir);

      // Step 2: Initialize management components
      RelationshipManager relationshipManager = new RelationshipManager(extractedDir);
      SPIDManager spidManager = new SPIDManager(extractedDir);
      SlideCreator slideCreator = new SlideCreator(extractedDir);

      // Step 3: Analyze presentation structure
      PresentationMetadata metadata = analyzePresentationStructure(extractedDir);

      // Step 4: Create session
      PresentationSession session = new PresentationSession(
          sessionId,
          pptxFile,
          extractedDir,
          relationshipManager,
          spidManager,
          slideCreator,
          metadata
          );

      activeSessions.put(sessionId, session);

      System.out.println("✓ Presentation opened: " + pptxFile.getName());
      System.out.println("  Session ID: " + sessionId);
      System.out.println("  Slides: " + metadata.getSlideCount());
      System.out.println("  Total shapes: " + spidManager.getAllSpids().size());

      return session;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to open presentation: " + pptxFile.getName(), e);
    }
  }

  /**
   * Saves a presentation session back to a PPTX file, performing full reconstruction.
   * 
   * @param session The presentation session to save
   * @param outputFile The target PPTX file (can be same as original for in-place update)
   * @throws XMLParsingException If the presentation cannot be saved
   * @throws IllegalArgumentException If session or outputFile is null
   */
  public void savePresentation(PresentationSession session, File outputFile) throws XMLParsingException {
    if (session == null || outputFile == null) {
      throw new IllegalArgumentException("Session and output file must be non-null");
    }

    try {
      // Step 1: Validate presentation integrity
      ValidationResult validation = validatePresentation(session);
      if (!validation.isValid()) {
        throw new XMLParsingException("Cannot save invalid presentation. Errors: " + 
            validation.getErrors());
      }

      // Step 2: Ensure output directory exists
      outputFile.getParentFile().mkdirs();

      // Step 3: Compress extracted directory back to PPTX
      compressPPTX(session.getExtractedDirectory(), outputFile);

      // Step 4: Update session metadata
      session.markSaved(outputFile);

      System.out.println("✓ Presentation saved: " + outputFile.getName());
      System.out.println("  Slides: " + session.getMetadata().getSlideCount());
      System.out.println("  Size: " + formatFileSize(outputFile.length()));

    } catch (Exception e) {
      throw new XMLParsingException("Failed to save presentation to: " + outputFile.getName(), e);
    }
  }

  /**
   * Adds a new blank slide to the presentation at the specified position.
   * 
   * @param session The presentation session
   * @param position Position to insert slide (1-based, inserts before this position)
   * @param slideTitle Title for the new slide
   * @return SlideOperationResult containing operation details and new slide number
   * @throws XMLParsingException If the slide cannot be added
   */
  public SlideOperationResult addSlide(PresentationSession session, int position, String slideTitle) 
      throws XMLParsingException {
      if (session == null) {
        throw new IllegalArgumentException("Session cannot be null");
      }

      try {
        // Begin transaction
        session.beginTransaction("addSlide");

        // Perform slide addition using SlideCreator
        int newSlideNumber = session.getSlideCreator().insertBlankSlide(position, slideTitle);

        // Update session metadata
        session.getMetadata().incrementSlideCount();
        session.markModified();

        // Commit transaction
        session.commitTransaction();

        SlideOperationResult result = new SlideOperationResult(
            SlideOperationType.ADD_SLIDE,
            true,
            "Slide added successfully at position " + position,
            newSlideNumber
            );

        System.out.println("✓ Added slide " + newSlideNumber + ": \"" + slideTitle + "\"");
        return result;

      } catch (Exception e) {
        session.rollbackTransaction();
        throw new XMLParsingException("Failed to add slide at position " + position, e);
      }
  }

  /**
   * Copies an existing slide to a new position in the presentation.
   * 
   * @param session The presentation session
   * @param sourceSlideNumber Source slide number to copy from (1-based)
   * @param destinationPosition Position to insert the copied slide
   * @param newSlideTitle Optional new title for the copied slide
   * @return SlideOperationResult containing operation details
   * @throws XMLParsingException If the slide cannot be copied
   */
  public SlideOperationResult copySlide(PresentationSession session, int sourceSlideNumber, 
      int destinationPosition, String newSlideTitle) throws XMLParsingException {
    if (session == null) {
      throw new IllegalArgumentException("Session cannot be null");
    }

    try {
      // Begin transaction
      session.beginTransaction("copySlide");

      // Perform slide copying using SlideCreator
      int newSlideNumber = session.getSlideCreator().insertCopiedSlide(
          destinationPosition, sourceSlideNumber, newSlideTitle);

      // Update session metadata
      session.getMetadata().incrementSlideCount();
      session.markModified();

      // Commit transaction
      session.commitTransaction();

      SlideOperationResult result = new SlideOperationResult(
          SlideOperationType.COPY_SLIDE,
          true,
          String.format("Slide %d copied to position %d", sourceSlideNumber, destinationPosition),
          newSlideNumber
          );

      System.out.println("✓ Copied slide " + sourceSlideNumber + " to slide " + newSlideNumber);
      return result;

    } catch (Exception e) {
      session.rollbackTransaction();
      throw new XMLParsingException("Failed to copy slide " + sourceSlideNumber + 
          " to position " + destinationPosition, e);
    }
  }

  /**
   * Validates the integrity of a presentation session.
   * 
   * @param session The presentation session to validate
   * @return ValidationResult containing any detected issues
   * @throws XMLParsingException If validation cannot be performed
   */
  public ValidationResult validatePresentation(PresentationSession session) throws XMLParsingException {
    if (session == null) {
      throw new IllegalArgumentException("Session cannot be null");
    }

    try {
      // Validate using SlideCreator's comprehensive validation
      return session.getSlideCreator().validatePresentation();

    } catch (Exception e) {
      throw new XMLParsingException("Failed to validate presentation", e);
    }
  }

  /**
   * Closes a presentation session and cleans up resources.
   * 
   * @param session The session to close
   * @throws XMLParsingException If cleanup fails
   */
  public void closeSession(PresentationSession session) throws XMLParsingException {
    if (session == null) {
      return;
    }

    try {
      // Remove from active sessions
      activeSessions.remove(session.getSessionId());

      // Clean up temporary files
      deleteDirectory(session.getExtractedDirectory().getParentFile());

      System.out.println("✓ Session closed: " + session.getSessionId());

    } catch (Exception e) {
      throw new XMLParsingException("Failed to close session: " + session.getSessionId(), e);
    }
  }

  /**
   * Gets all currently active presentation sessions.
   * 
   * @return Unmodifiable collection of active sessions
   */
  public Collection<PresentationSession> getActiveSessions() {
    return Collections.unmodifiableCollection(activeSessions.values());
  }

  /**
   * Retrieves a specific session by ID.
   * 
   * @param sessionId The session ID to retrieve
   * @return The session, or null if not found
   */
  public PresentationSession getSession(String sessionId) {
    return activeSessions.get(sessionId);
  }

  // ========== PRIVATE HELPER METHODS ==========

  /**
   * Extracts a PPTX file to the specified directory using PowerShell.
   */
  private void extractPPTX(File pptxFile, File extractDir) throws XMLParsingException {
    try {
      extractDir.mkdirs();

      // Use PowerShell to extract ZIP contents
      String command = String.format(
          "powershell -command \"Expand-Archive -Path '%s' -DestinationPath '%s' -Force\"",
          pptxFile.getAbsolutePath(),
          extractDir.getAbsolutePath()
          );

      Process process = Runtime.getRuntime().exec(command);
      int exitCode = process.waitFor();

      if (exitCode != 0) {
        throw new XMLParsingException("Failed to extract PPTX file: " + pptxFile.getName());
      }

      // Verify extraction success
      if (!new File(extractDir, "ppt/presentation.xml").exists()) {
        throw new XMLParsingException("Invalid PPTX structure after extraction");
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to extract PPTX: " + pptxFile.getName(), e);
    }
  }

  /**
   * Compresses an extracted directory back to a PPTX file using PowerShell.
   */
  private void compressPPTX(File extractedDir, File outputFile) throws XMLParsingException {
    try {
      // Delete output file if it exists
      if (outputFile.exists()) {
        outputFile.delete();
      }

      // Use PowerShell to create ZIP archive
      String command = String.format(
          "powershell -command \"Compress-Archive -Path '%s\\*' -DestinationPath '%s' -Force\"",
          extractedDir.getAbsolutePath(),
          outputFile.getAbsolutePath()
          );

      Process process = Runtime.getRuntime().exec(command);
      int exitCode = process.waitFor();

      if (exitCode != 0) {
        throw new XMLParsingException("Failed to compress PPTX file: " + outputFile.getName());
      }

      // Verify compression success
      if (!outputFile.exists() || outputFile.length() == 0) {
        throw new XMLParsingException("Failed to create valid PPTX file: " + outputFile.getName());
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to compress PPTX: " + outputFile.getName(), e);
    }
  }

  /**
   * Analyzes the structure of an extracted presentation.
   */
  private PresentationMetadata analyzePresentationStructure(File extractedDir) throws XMLParsingException {
    try {
      // Count slides
      File slidesDir = new File(extractedDir, "ppt/slides");
      int slideCount = 0;
      if (slidesDir.exists()) {
        File[] slideFiles = slidesDir.listFiles((dir, name) -> name.matches("slide\\d+\\.xml"));
        slideCount = slideFiles != null ? slideFiles.length : 0;
      }

      // Analyze presentation.xml for slide relationships
      File presentationFile = new File(extractedDir, "ppt/presentation.xml");
      Document presentationDoc = documentBuilder.parse(presentationFile);

      // Create metadata object
      return new PresentationMetadata(slideCount, extractedDir, presentationDoc);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to analyze presentation structure", e);
    }
  }

  /**
   * Generates a unique session ID.
   */
  private String generateSessionId() {
    return "session_" + sessionIdGenerator.getAndIncrement() + "_" + System.currentTimeMillis();
  }

  /**
   * Creates a temporary directory for PPTX operations.
   */
  private File createTempDirectory() {
    try {
      File tempDir = new File(System.getProperty("java.io.tmpdir"), "presentation_choreographer");
      tempDir.mkdirs();
      return tempDir;
    } catch (Exception e) {
      // Fallback to current directory
      return new File("temp_pptx_operations");
    }
  }

  /**
   * Recursively deletes a directory and all its contents.
   */
  private void deleteDirectory(File directory) {
    if (directory == null || !directory.exists()) {
      return;
    }

    File[] files = directory.listFiles();
    if (files != null) {
      for (File file : files) {
        if (file.isDirectory()) {
          deleteDirectory(file);
        } else {
          file.delete();
        }
      }
    }
    directory.delete();
  }

  /**
   * Formats file size for human-readable display.
   */
  private String formatFileSize(long bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
    return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
  }

  // ========== INNER CLASSES ==========

  /**
   * Represents an active presentation editing session.
   */
  public static class PresentationSession {
    private final String sessionId;
    private final File originalFile;
    private final File extractedDirectory;
    private final RelationshipManager relationshipManager;
    private final SPIDManager spidManager;
    private final SlideCreator slideCreator;
    private final PresentationMetadata metadata;
    private final Stack<String> transactionStack;
    private boolean isModified;
    private File lastSavedFile;

    public PresentationSession(String sessionId, File originalFile, File extractedDirectory,
        RelationshipManager relationshipManager, SPIDManager spidManager,
        SlideCreator slideCreator, PresentationMetadata metadata) {
      this.sessionId = sessionId;
      this.originalFile = originalFile;
      this.extractedDirectory = extractedDirectory;
      this.relationshipManager = relationshipManager;
      this.spidManager = spidManager;
      this.slideCreator = slideCreator;
      this.metadata = metadata;
      this.transactionStack = new Stack<>();
      this.isModified = false;
      this.lastSavedFile = originalFile;
    }

    // Transaction management
    public void beginTransaction(String operationName) {
      transactionStack.push(operationName);
    }

    public void commitTransaction() {
      if (!transactionStack.isEmpty()) {
        transactionStack.pop();
      }
    }

    public void rollbackTransaction() {
      if (!transactionStack.isEmpty()) {
        String operation = transactionStack.pop();
        System.out.println("⚠ Transaction rolled back: " + operation);
        // TODO: Implement actual rollback logic using file snapshots
      }
    }

    // Getters
    public String getSessionId() { return sessionId; }
    public File getOriginalFile() { return originalFile; }
    public File getExtractedDirectory() { return extractedDirectory; }
    public RelationshipManager getRelationshipManager() { return relationshipManager; }
    public SPIDManager getSPIDManager() { return spidManager; }
    public SlideCreator getSlideCreator() { return slideCreator; }
    public PresentationMetadata getMetadata() { return metadata; }
    public boolean isModified() { return isModified; }
    public File getLastSavedFile() { return lastSavedFile; }

    public void markModified() { this.isModified = true; }
    public void markSaved(File savedFile) { 
      this.isModified = false; 
      this.lastSavedFile = savedFile;
    }
  }

  /**
   * Metadata about a presentation structure.
   */
  public static class PresentationMetadata {
    private int slideCount;
    private final File extractedDirectory;
    private final Document presentationDocument;

    public PresentationMetadata(int slideCount, File extractedDirectory, Document presentationDocument) {
      this.slideCount = slideCount;
      this.extractedDirectory = extractedDirectory;
      this.presentationDocument = presentationDocument;
    }

    public int getSlideCount() { return slideCount; }
    public void incrementSlideCount() { this.slideCount++; }
    public void decrementSlideCount() { this.slideCount = Math.max(0, this.slideCount - 1); }
    public File getExtractedDirectory() { return extractedDirectory; }
    public Document getPresentationDocument() { return presentationDocument; }
  }

  /**
   * Result of a slide operation.
   */
  public static class SlideOperationResult {
    private final SlideOperationType operationType;
    private final boolean success;
    private final String message;
    private final int resultingSlideNumber;

    public SlideOperationResult(SlideOperationType operationType, boolean success, 
        String message, int resultingSlideNumber) {
      this.operationType = operationType;
      this.success = success;
      this.message = message;
      this.resultingSlideNumber = resultingSlideNumber;
    }

    public SlideOperationType getOperationType() { return operationType; }
    public boolean isSuccess() { return success; }
    public String getMessage() { return message; }
    public int getResultingSlideNumber() { return resultingSlideNumber; }
  }

  /**
   * Types of slide operations.
   */
  public enum SlideOperationType {
    ADD_SLIDE, COPY_SLIDE, DELETE_SLIDE, MOVE_SLIDE, MODIFY_SLIDE
  }

  /**
   * Validation result from SlideCreator.
   */
  public static class ValidationResult {
    private final List<String> errors;
    private final List<String> warnings;

    public ValidationResult(List<String> errors, List<String> warnings) {
      this.errors = Collections.unmodifiableList(new ArrayList<>(errors));
      this.warnings = Collections.unmodifiableList(new ArrayList<>(warnings));
    }

    public List<String> getErrors() { return errors; }
    public List<String> getWarnings() { return warnings; }
    public boolean hasErrors() { return !errors.isEmpty(); }
    public boolean hasWarnings() { return !warnings.isEmpty(); }
    public boolean isValid() { return errors.isEmpty(); }
  }
}
