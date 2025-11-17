/**
 * Abstract base class for all controllers
 * Provides reusable pattern for event handler management
 * All controllers should extend this class
 */
export abstract class BaseController {
  protected eventListeners: Array<{
    element: EventTarget;
    event: string;
    handler: EventListener;
  }> = [];

  /**
   * Initialize the controller
   * Must be implemented by subclasses
   */
  abstract initialize(): void;

  /**
   * Add event listener with automatic cleanup tracking
   */
  protected addEventListener(
    element: EventTarget,
    event: string,
    handler: EventListener
  ): void {
    element.addEventListener(event, handler);
    this.eventListeners.push({ element, event, handler });
  }

  /**
   * Remove all event listeners and clean up
   */
  destroy(): void {
    this.eventListeners.forEach(({ element, event, handler }) => {
      element.removeEventListener(event, handler);
    });
    this.eventListeners = [];
  }
}

