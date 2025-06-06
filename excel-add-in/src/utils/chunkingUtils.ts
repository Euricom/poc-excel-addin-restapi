/**
 * Utility functions for implementing chunking strategies with Excel API
 * to prevent exceeding the 5MB payload limit and optimize performance.
 */

import { timeAsync } from './performanceUtils'

// Constants for chunking configuration
export const CHUNKING_CONFIG = {
  // Default chunk sizes (adjust based on testing)
  DEFAULT_CHUNK_SIZE: 250,
  MIN_CHUNK_SIZE: 50,
  MAX_CHUNK_SIZE: 500,

  // Retry configuration
  MAX_RETRIES: 3,
  INITIAL_RETRY_DELAY: 1000, // 1 second
  MAX_RETRY_DELAY: 10000, // 10 seconds

  // Progress tracking
  PROGRESS_KEY_PREFIX: 'excel_chunking_progress_',

  // Payload size estimation
  ESTIMATED_BYTES_PER_CELL: 50, // Conservative estimate
  MAX_PAYLOAD_SIZE_BYTES: 5 * 1024 * 1024 // 5MB
}

/**
 * Interface for chunking progress tracking
 */
export interface ChunkingProgress {
  operationId: string
  totalChunks: number
  completedChunks: number
  startTime: number
  lastChunkTime: number
  status: 'in_progress' | 'completed' | 'failed'
  error?: string
}

/**
 * Interface for chunking options
 */
export interface ChunkingOptions {
  chunkSize?: number
  maxRetries?: number
  operationId?: string
  onProgress?: (progress: ChunkingProgress) => void
  resumeFromChunk?: number
}

/**
 * Estimates the payload size in bytes for a range of values
 * @param values The 2D array of values
 * @returns Estimated size in bytes
 */
export function estimatePayloadSize(values: any[][]): number {
  if (!values || values.length === 0) return 0

  let totalCells = 0

  // Count total cells
  for (let i = 0; i < values.length; i++) {
    totalCells += values[i].length
  }

  // Estimate size based on average bytes per cell
  return totalCells * CHUNKING_CONFIG.ESTIMATED_BYTES_PER_CELL
}

/**
 * Determines optimal chunk size based on data characteristics
 * @param totalRows Total number of rows
 * @param columnsPerRow Number of columns per row
 * @returns Optimal chunk size (number of rows per chunk)
 */
export function calculateOptimalChunkSize(
  totalRows: number,
  columnsPerRow: number
): number {
  // Calculate cells per chunk based on max payload size
  const maxCellsPerChunk =
    CHUNKING_CONFIG.MAX_PAYLOAD_SIZE_BYTES /
    CHUNKING_CONFIG.ESTIMATED_BYTES_PER_CELL

  // Calculate rows per chunk
  const rowsPerChunk = Math.floor(maxCellsPerChunk / columnsPerRow)

  // Constrain to configured limits
  return Math.min(
    Math.max(rowsPerChunk, CHUNKING_CONFIG.MIN_CHUNK_SIZE),
    CHUNKING_CONFIG.MAX_CHUNK_SIZE
  )
}

/**
 * Splits a 2D array of values into chunks of specified size
 * @param values The 2D array to chunk
 * @param chunkSize Number of rows per chunk
 * @returns Array of chunked 2D arrays
 */
export function chunkRangeValues(
  values: any[][],
  chunkSize: number
): any[][][] {
  const chunks: any[][][] = []

  for (let i = 0; i < values.length; i += chunkSize) {
    chunks.push(values.slice(i, i + chunkSize))
  }

  return chunks
}

/**
 * Creates a new progress tracker for a chunking operation
 * @param operationId Unique identifier for the operation
 * @param totalChunks Total number of chunks
 * @returns Progress object
 */
export function createProgressTracker(
  operationId: string,
  totalChunks: number
): ChunkingProgress {
  const progress: ChunkingProgress = {
    operationId,
    totalChunks,
    completedChunks: 0,
    startTime: Date.now(),
    lastChunkTime: Date.now(),
    status: 'in_progress'
  }

  // Save to session storage for persistence
  saveProgress(progress)

  return progress
}

/**
 * Updates progress for a chunking operation
 * @param progress Progress object to update
 * @param completedChunks Number of completed chunks
 * @param status Optional status update
 * @param error Optional error message
 * @returns Updated progress object
 */
export function updateProgress(
  progress: ChunkingProgress,
  completedChunks: number,
  status?: 'in_progress' | 'completed' | 'failed',
  error?: string
): ChunkingProgress {
  progress.completedChunks = completedChunks
  progress.lastChunkTime = Date.now()

  if (status) {
    progress.status = status
  }

  if (error) {
    progress.error = error
  }

  // Save updated progress
  saveProgress(progress)

  return progress
}

/**
 * Saves progress to session storage
 * @param progress Progress object to save
 */
function saveProgress(progress: ChunkingProgress): void {
  try {
    const key = `${CHUNKING_CONFIG.PROGRESS_KEY_PREFIX}${progress.operationId}`
    sessionStorage.setItem(key, JSON.stringify(progress))
  } catch (error) {
    console.warn('Failed to save chunking progress to session storage:', error)
  }
}

/**
 * Loads progress from session storage
 * @param operationId Operation identifier
 * @returns Progress object or null if not found
 */
export function loadProgress(operationId: string): ChunkingProgress | null {
  try {
    const key = `${CHUNKING_CONFIG.PROGRESS_KEY_PREFIX}${operationId}`
    const saved = sessionStorage.getItem(key)

    if (saved) {
      return JSON.parse(saved) as ChunkingProgress
    }
  } catch (error) {
    console.warn(
      'Failed to load chunking progress from session storage:',
      error
    )
  }

  return null
}

/**
 * Clears progress from session storage
 * @param operationId Operation identifier
 */
export function clearProgress(operationId: string): void {
  try {
    const key = `${CHUNKING_CONFIG.PROGRESS_KEY_PREFIX}${operationId}`
    sessionStorage.removeItem(key)
  } catch (error) {
    console.warn(
      'Failed to clear chunking progress from session storage:',
      error
    )
  }
}

/**
 * Implements exponential backoff for retrying operations
 * @param operation Function to retry
 * @param maxRetries Maximum number of retry attempts
 * @param initialDelay Initial delay in milliseconds
 * @param maxDelay Maximum delay in milliseconds
 * @returns Promise resolving to the operation result
 */
export async function withExponentialBackoff<T>(
  operation: () => Promise<T>,
  maxRetries: number = CHUNKING_CONFIG.MAX_RETRIES,
  initialDelay: number = CHUNKING_CONFIG.INITIAL_RETRY_DELAY,
  maxDelay: number = CHUNKING_CONFIG.MAX_RETRY_DELAY
): Promise<T> {
  let retries = 0
  let delay = initialDelay

  while (true) {
    try {
      return await operation()
    } catch (error) {
      retries++

      // If we've reached max retries, throw the error
      if (retries > maxRetries) {
        throw error
      }

      console.warn(
        `Operation failed, retrying (${retries}/${maxRetries}) after ${delay}ms:`,
        error
      )

      // Wait for the backoff delay
      await new Promise(resolve => setTimeout(resolve, delay))

      // Exponential backoff with jitter
      delay = Math.min(delay * 2 * (0.9 + Math.random() * 0.2), maxDelay)
    }
  }
}

/**
 * Processes a large range operation in chunks with Excel API
 * @param context Excel RequestContext
 * @param sheet Target worksheet
 * @param values 2D array of values to write
 * @param startCell Starting cell reference (e.g., "A1")
 * @param options Chunking options
 * @returns Promise resolving when all chunks are processed
 */
export async function processRangeInChunks(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  values: any[][],
  startCell: string,
  options: ChunkingOptions = {}
): Promise<void> {
  // Extract options with defaults
  const {
    chunkSize = CHUNKING_CONFIG.DEFAULT_CHUNK_SIZE,
    maxRetries = CHUNKING_CONFIG.MAX_RETRIES,
    operationId = `range_op_${Date.now()}`,
    onProgress,
    resumeFromChunk = 0
  } = options

  // If values array is empty, nothing to do
  if (!values || values.length === 0) return

  // Calculate optimal chunk size if not specified
  const actualChunkSize =
    chunkSize || calculateOptimalChunkSize(values.length, values[0].length)

  // Split values into chunks
  const chunks = chunkRangeValues(values, actualChunkSize)

  // Create or load progress tracker
  let progress =
    loadProgress(operationId) ||
    createProgressTracker(operationId, chunks.length)

  // Update progress with current chunk count (in case it changed)
  progress.totalChunks = chunks.length

  // Skip already completed chunks if resuming
  const startChunk = Math.max(0, Math.min(resumeFromChunk, chunks.length - 1))

  // Process each chunk
  for (let i = startChunk; i < chunks.length; i++) {
    const chunk = chunks[i]
    const chunkStartRow = i * actualChunkSize + 1 // 1-based row index

    // Parse the starting cell to get column
    const colMatch = startCell.match(/([A-Z]+)/)
    const startColumn = colMatch ? colMatch[0] : 'A'

    // Calculate range address for this chunk
    const chunkRange = `${startColumn}${chunkStartRow}:${getColumnLetter(
      startColumn,
      chunk[0].length - 1
    )}${chunkStartRow + chunk.length - 1}`

    // Process this chunk with retry logic
    await timeAsync(
      `processChunk-${i + 1}-of-${chunks.length}`,
      async () => {
        await withExponentialBackoff(async () => {
          // Get the range and set values
          const range = sheet.getRange(chunkRange)
          range.values = chunk

          // Sync after each chunk
          await context.sync()
        }, maxRetries)
      },
      { chunkSize: chunk.length, rowCount: chunk.length * chunk[0].length }
    )

    // Update progress
    progress = updateProgress(progress, i + 1)

    // Call progress callback if provided
    if (onProgress) {
      onProgress(progress)
    }
  }

  // Mark operation as completed
  updateProgress(progress, chunks.length, 'completed')

  // Clear progress after successful completion
  clearProgress(operationId)
}

/**
 * Helper function to get Excel column letter from index
 * @param startColumn Starting column letter
 * @param columnOffset Number of columns to offset
 * @returns Column letter
 */
function getColumnLetter(startColumn: string, columnOffset: number): string {
  if (columnOffset === 0) return startColumn

  // Convert column letters to number (A=1, B=2, etc.)
  let num = 0
  for (let i = 0; i < startColumn.length; i++) {
    num = num * 26 + (startColumn.charCodeAt(i) - 64)
  }

  // Add offset
  num += columnOffset

  // Convert back to letters
  let result = ''
  while (num > 0) {
    const remainder = (num - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    num = Math.floor((num - 1) / 26)
  }

  return result
}

/**
 * Processes a large dataset with chunking, retries, and progress tracking
 * @param processChunkFn Function to process each chunk
 * @param items Array of items to process
 * @param options Chunking options
 * @returns Promise resolving when all chunks are processed
 */
export async function processItemsInChunks<T>(
  processChunkFn: (
    chunk: T[],
    chunkIndex: number,
    totalChunks: number
  ) => Promise<void>,
  items: T[],
  options: ChunkingOptions = {}
): Promise<void> {
  // Extract options with defaults
  const {
    chunkSize = CHUNKING_CONFIG.DEFAULT_CHUNK_SIZE,
    maxRetries = CHUNKING_CONFIG.MAX_RETRIES,
    operationId = `items_op_${Date.now()}`,
    onProgress,
    resumeFromChunk = 0
  } = options

  // If items array is empty, nothing to do
  if (!items || items.length === 0) return

  // Split items into chunks
  const chunks: T[][] = []
  for (let i = 0; i < items.length; i += chunkSize) {
    chunks.push(items.slice(i, i + chunkSize))
  }

  // Create or load progress tracker
  let progress =
    loadProgress(operationId) ||
    createProgressTracker(operationId, chunks.length)

  // Update progress with current chunk count (in case it changed)
  progress.totalChunks = chunks.length

  // Skip already completed chunks if resuming
  const startChunk = Math.max(0, Math.min(resumeFromChunk, chunks.length - 1))

  // Process each chunk
  for (let i = startChunk; i < chunks.length; i++) {
    const chunk = chunks[i]

    // Process this chunk with retry logic
    await timeAsync(
      `processItemsChunk-${i + 1}-of-${chunks.length}`,
      async () => {
        await withExponentialBackoff(async () => {
          await processChunkFn(chunk, i, chunks.length)
        }, maxRetries)
      },
      { chunkSize: chunk.length }
    )

    // Update progress
    progress = updateProgress(progress, i + 1)

    // Call progress callback if provided
    if (onProgress) {
      onProgress(progress)
    }
  }

  // Mark operation as completed
  updateProgress(progress, chunks.length, 'completed')

  // Clear progress after successful completion
  clearProgress(operationId)
}
