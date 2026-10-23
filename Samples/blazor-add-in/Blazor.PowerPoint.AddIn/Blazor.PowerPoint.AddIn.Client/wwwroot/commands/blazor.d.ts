/**
 * Type definitions for Blazor's DotNet interop
 */

/**
 * Represents a DotNetObjectReference received from .NET.
 * Methods are invoked on the specific .NET object instance,
 * regardless of render mode (InteractiveServer or InteractiveWebAssembly).
 */
declare namespace DotNet {
  interface DotNetObject {
    invokeMethodAsync<T>(methodName: string, ...args: any[]): Promise<T>;
    invokeMethod<T>(methodName: string, ...args: any[]): T;
    dispose(): void;
  }
}

/**
 * Global DotNet object for static method invocation (assembly-based).
 * Prefer DotNet.DotNetObject for instance-based invocation via DotNetObjectReference.
 */
interface DotNet {
  invokeMethodAsync<T>(assemblyName: string, methodName: string, ...args: any[]): Promise<T>;
  invokeMethod<T>(assemblyName: string, methodName: string, ...args: any[]): T;
}

declare const DotNet: DotNet;

/**
 * Blazor interop properties exposed on window by the inline script in App.razor.
 * Augments the global Window interface for type-safe access.
 */
interface Window {
  dotNetReady: { [bridgeName: string]: Promise<void> };
  dotNetRefs: Map<string, DotNet.DotNetObject>;
  signalDotNetReady(name: string, dotNetRef: DotNet.DotNetObject): void;
}
