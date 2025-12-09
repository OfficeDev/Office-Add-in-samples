/**
 * Type definitions for Blazor's DotNet global object
 * 
 * The global DotNet object is injected by the Blazor WebAssembly runtime at runtime, 
 * not during compilation. It's part of the Blazor WebAssembly JavaScript interop system 
 * but isn't formally defined in the TypeScript type system provided by Microsoft.
 * 
 * This is why many developers end up creating their own type definitions for the DotNet object, 
 * as Microsoft doesn't currently provide an "official" TypeScript declaration for it in their 
 * packages.
 */
interface DotNet {
  invokeMethodAsync<T>(assemblyName: string, methodName: string, ...args: any[]): Promise<T>;
  invokeMethod<T>(assemblyName: string, methodName: string, ...args: any[]): T;
}

declare const DotNet: DotNet;