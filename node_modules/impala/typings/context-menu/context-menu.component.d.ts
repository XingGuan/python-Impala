import { EventEmitter, OnInit } from '@angular/core';
import { ContextMenuConfig } from "./context-menu-config";
import { ContextMenuAction } from "./context-menu-item";
export declare class ContextMenuComponent implements OnInit {
    config: ContextMenuConfig;
    menuActionClick: EventEmitter<{}>;
    constructor(config: ContextMenuConfig);
    ngOnInit(): void;
    onContextMenuAction(item: ContextMenuAction): void;
}
