/**
 * Type declarations for shared Office utility functions exposed globally by commands.ts.
 */
declare function goToLastSlide(): Promise<void>;
declare function insertImage(base64Image: string): Promise<void>;
declare function removeSlidePlaceholders(shapes: PowerPoint.ShapeCollection): Promise<void>;
