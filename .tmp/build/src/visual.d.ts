import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import visualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import "./../style/visual.less";
export declare class Visual implements IVisual {
    private svg;
    private margin;
    private barStep;
    private barPadding;
    private duration;
    private color;
    private x;
    private width;
    private height;
    private xAxis;
    private yAxis;
    constructor(options: VisualConstructorOptions);
    update(options: visualUpdateOptions): void;
    private bar;
    private down;
    private stack;
    private stagger;
    private up;
}
