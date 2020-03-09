import { ContextMenuConfig } from "./context-menu-config";
import { ContextMenuItem } from "./context-menu-item";
export declare class ContextMenuConfigBuilder {
    config: ContextMenuConfig;
    constructor();
    left(pageX: number): ContextMenuConfigBuilder;
    top(pageY: number): ContextMenuConfigBuilder;
    addAction(text: string, click: (context?: any) => any, disabled?: boolean): ContextMenuConfigBuilder;
    addDivider(): ContextMenuConfigBuilder;
    addItems(...items: ContextMenuItem[]): ContextMenuConfigBuilder;
    context(context: any): this;
    build(): ContextMenuConfig;
}
