/**
 * Utility functions for performance monitoring and optimization
 */

// Performance metrics storage
interface PerformanceMetric {
  operation: string
  startTime: number
  endTime: number
  duration: number
  metadata?: Record<string, any>
}

const performanceMetrics: PerformanceMetric[] = []
const MAX_METRICS = 100 // Maximum number of metrics to store

/**
 * Start timing an operation
 * @param operation Name of the operation being timed
 * @param metadata Optional metadata about the operation
 * @returns A unique identifier for this timing operation
 */
export function startTiming(
  operation: string,
  metadata?: Record<string, any>
): number {
  // Clean up old metrics if we're at capacity
  if (performanceMetrics.length >= MAX_METRICS) {
    performanceMetrics.splice(0, Math.floor(MAX_METRICS / 2))
  }

  const metric: PerformanceMetric = {
    operation,
    startTime: performance.now(),
    endTime: 0,
    duration: 0,
    metadata
  }

  performanceMetrics.push(metric)
  return performanceMetrics.length - 1
}

/**
 * End timing an operation and record its duration
 * @param id The identifier returned from startTiming
 * @returns The duration of the operation in milliseconds
 */
export function endTiming(id: number): number {
  if (id < 0 || id >= performanceMetrics.length) {
    console.error(`Invalid timing ID: ${id}`)
    return 0
  }

  const metric = performanceMetrics[id]
  metric.endTime = performance.now()
  metric.duration = metric.endTime - metric.startTime

  console.log(
    `Performance: ${metric.operation} took ${metric.duration.toFixed(2)}ms`,
    metric.metadata || ''
  )

  return metric.duration
}

/**
 * Wraps an async function with performance timing
 * @param operation Name of the operation
 * @param fn The async function to time
 * @param metadata Optional metadata
 * @returns The result of the function
 */
export async function timeAsync<T>(
  operation: string,
  fn: () => Promise<T>,
  metadata?: Record<string, any>
): Promise<T> {
  const timerId = startTiming(operation, metadata)
  try {
    return await fn()
  } finally {
    endTiming(timerId)
  }
}

/**
 * Get performance metrics for analysis
 * @param operation Optional filter by operation name
 * @returns Array of performance metrics
 */
export function getPerformanceMetrics(operation?: string): PerformanceMetric[] {
  if (operation) {
    return performanceMetrics.filter(m => m.operation === operation)
  }
  return [...performanceMetrics]
}

/**
 * Calculate average duration for a specific operation
 * @param operation The operation name to analyze
 * @returns Average duration in milliseconds
 */
export function getAverageDuration(operation: string): number {
  const metrics = performanceMetrics.filter(
    m => m.operation === operation && m.duration > 0
  )

  if (metrics.length === 0) return 0

  const total = metrics.reduce((sum, metric) => sum + metric.duration, 0)
  return total / metrics.length
}

/**
 * Clear all performance metrics
 */
export function clearPerformanceMetrics(): void {
  performanceMetrics.length = 0
}

/**
 * Log a performance report to the console
 */
export function logPerformanceReport(): void {
  // Group metrics by operation
  const operationMap = new Map<string, number[]>()

  performanceMetrics.forEach(metric => {
    if (metric.duration > 0) {
      const durations = operationMap.get(metric.operation) || []
      durations.push(metric.duration)
      operationMap.set(metric.operation, durations)
    }
  })

  console.group('Performance Report')

  operationMap.forEach((durations, operation) => {
    const count = durations.length
    const total = durations.reduce((sum, d) => sum + d, 0)
    const avg = total / count
    const min = Math.min(...durations)
    const max = Math.max(...durations)

    console.log(
      `${operation}: ${count} calls, avg: ${avg.toFixed(
        2
      )}ms, min: ${min.toFixed(2)}ms, max: ${max.toFixed(2)}ms`
    )
  })

  console.groupEnd()
}
