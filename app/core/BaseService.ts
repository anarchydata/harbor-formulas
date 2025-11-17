/**
 * Abstract base class for all services
 * Provides reusable pattern for initialization and cleanup
 * All services should extend this class
 */
export abstract class BaseService {
  protected initialized = false;

  /**
   * Initialize the service
   * Must be implemented by subclasses
   */
  abstract initialize(): Promise<void> | void;

  /**
   * Clean up resources
   * Must be implemented by subclasses
   */
  abstract destroy(): void;

  /**
   * Check if service is initialized
   */
  isInitialized(): boolean {
    return this.initialized;
  }
}

