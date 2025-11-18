/**
 * Reusable service and controller interfaces
 * These can be used for type checking and dependency injection
 */

export interface IService {
  initialize(): Promise<void> | void;
  destroy(): void;
  isInitialized(): boolean;
}

export interface IController {
  initialize(): void;
  destroy(): void;
}

