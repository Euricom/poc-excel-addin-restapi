/**
 * Update status message in the UI
 */
export function updateStatus(message: string, isError: boolean = false): void {
  const statusElement = document.getElementById('status')
  if (statusElement) {
    statusElement.textContent = message
    statusElement.className = isError ? 'status error' : 'status success'
  }
}
