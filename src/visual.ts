import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import visualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import * as d3 from "d3";
import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { BasicFilter } from "powerbi-models";


// @ts-ignore: Allow side-effect import of LESS stylesheet without module declarations.
import "./../style/visual.less";

interface DataNode {
    name: string,
    value?: number,
    children?: DataNode[],
}

type Node = d3.HierarchyNode<DataNode> & { index: number, value: number };

function convertMatrixNode(
    node: powerbi.DataViewMatrixNode,
    host: IVisualHost
): DataNode {
    const result: DataNode = {
        name: node.value != null ? String(node.value) : "",
    };
    if (node.values) {
        result.value = (node.values[0] as any)?.value;
    }
    if (node.children) {
        result.children = node.children.map(child =>
            convertMatrixNode(child, host)
        );
    }
    return result;
}

export class Visual implements IVisual {
    private host!: IVisualHost;
    private matrix!: powerbi.DataViewMatrix | undefined;
    private levelFilters: Map<number, Set<string>> = new Map();

    private getFilterTarget(level: number): { table: string, column: string } {
        const source = this.matrix!.rows.levels[level].sources[0];
        const queryName = source.queryName; // "TableName.ColumnName"
        const lastDot = queryName!.lastIndexOf('.');
        return {
            table: queryName!.substring(0, lastDot),
            column: queryName!.substring(lastDot + 1)
        };
    }

    private applyCurrentFilters() {
        if (this.levelFilters.size === 0) {
            this.host.applyJsonFilter(
                [], "general", "filter", powerbi.FilterAction.remove
            );
            return;
        }

        const filters = Array.from(this.levelFilters.entries())
            .sort(([a], [b]) => a - b) // ensure level order
            .map(([level, values]) => new BasicFilter(
                this.getFilterTarget(level),
                "In",
                Array.from(values)
            ));

        this.host.applyJsonFilter(
            filters, "general", "filter", powerbi.FilterAction.merge
        );
    }

    private svg!: d3.Selection<SVGSVGElement, unknown, null, undefined>;

    private margin = { top: 30, right: 30, bottom: 0, left: 100 };
    private barStep = 27;
    private barPadding = 3 / this.barStep;
    private duration = 750;
    private isTransitioning = false;
    private color = d3.scaleOrdinal([true, false], ["steelblue", "#aaa"])
    
    private x!: d3.ScaleLinear<number, number>; // top scale
    private width = 0;
    private height = 0;

    private xAxis = (g: any) => {
        g.attr("class", "x-axis")
            .attr("transform", `translate(0,${this.margin.top})`)
            .call(d3.axisTop(this.x).ticks(this.width / 80, "s"))
            .call((g: any) => (g.selection ? g.selection() : g).select(".domain").remove());
        return g;
    }

    private yAxis = (g: any) => {
        g.attr("class", "y-axis")
            .attr("transform", `translate(${this.margin.left},0)`)
            .call((g: any) => g.append("line")
                .attr("stroke", "currentColor")
                .attr("y1", this.margin.top)
                .attr("y2", this.height - this.margin.bottom));
        return g;
    }


    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.svg = d3.select(options.element).append("svg")
    }

    public update(options: visualUpdateOptions) {
        this.matrix = options.dataViews[0].matrix;
        if (!this.matrix?.rows?.root?.children?.length) return;

        const data: DataNode = {
            name: "root",
            children: this.matrix.rows.root.children.map(child =>
                convertMatrixNode(child, this.host)
            )
        };
        
        const root = d3.hierarchy<DataNode>(data)
            .sum(d => d.value ?? 0)
            .sort((a, b) => (b.value ?? 0) - (a.value ?? 0)) as Node;

        root.eachAfter(d => (d as Node).index = d.parent
            ? (d.parent as Node).index = ((d.parent as Node).index + 1 || 0)
            : 0);

        this.width = options.viewport.width;
        let max = 0;
        root.each(d => d.children && (max = Math.max(max, d.children.length)));
        this.height = max * this.barStep + this.margin.top + this.margin.bottom;

        this.x = d3.scaleLinear().range([this.margin.left, this.width - this.margin.right]);
        this.x.domain([0, root.value ?? 0]);

        this.svg.interrupt();
        this.svg.selectAll("*").remove();

        this.svg
            .attr("viewBox", `0 0 ${this.width} ${this.height}`)
            .attr("width", `${this.width}`)
            .attr("height", `${this.height}`)
            .attr("style", "max-width: 100%; height: auto;")

        this.svg
            .append("rect")
            .attr("class", "background")
            .attr("fill", "none")
            .attr("pointer-events", "all")
            .attr("width", this.width)
            .attr("height", this.height)
            .attr("cursor", "pointer")
            .on("click", (_event: MouseEvent, d: any) => {
                if (this.isTransitioning) return;
                if (!d.parent || d.parent.depth === 0) {
                    this.levelFilters.clear();
                } else {
                    const parentAncestors = d.parent.ancestors().reverse().slice(1);
                    this.levelFilters.clear();
                    parentAncestors.forEach((ancestor: any, i: number) => {
                        this.levelFilters.set(i, new Set([ancestor.data.name]));
                    });
                }

                this.applyCurrentFilters();
                this.up(d);
            });

        this.svg
            .append("g")
            .call(this.xAxis);
        this.svg
            .append("g")
            .call(this.yAxis);

        this.down(root);
    }

    private bar(d: Node, selector: string) {
        const g = this.svg.insert("g", selector)
            .attr("class", "enter")
            .attr("transform", `translate(0,${this.margin.top + this.barStep * this.barPadding})`)
            .attr("text-anchor", "end")
            .style("font", "10px sans-serif");

        const bar = g.selectAll("g")
            .data(d.children ? d.children : [])
            .join("g")
            .attr("cursor", d => !d.children ? null : "pointer")
            .on("click", (event, d) => {
                if (this.isTransitioning) return;

                const isLeaf = !d.children || d.children.length === 0;
                const ctrlKey = (event as MouseEvent).ctrlKey;

                const ancestors = d.ancestors().reverse().slice(1);
                if (isLeaf && ctrlKey) {
                    const leafLevel = d.depth - 1;
                    if (!this.levelFilters.has(leafLevel)) {
                        this.levelFilters.set(leafLevel, new Set());
                    }
                    this.levelFilters.get(leafLevel)!.add(d.data.name);
                } else {

                    this.levelFilters.clear();
                    ancestors.forEach((ancestor, i) => {
                        this.levelFilters.set(i, new Set([ancestor.data.name]));
                    });
                }
                this.applyCurrentFilters();
                this.down(d);
            });

        bar.append("text")
            .attr("x", this.margin.left - 6)
            .attr("y", this.barStep * (1 - this.barPadding) / 2)
            .attr("dy", ".35em")
            .text(d => d.data.name);

        bar.append("rect")
            .attr("x", this.x(0))
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .attr("height", this.barStep * (1 - this.barPadding));
        return g;
    }

    private down(d: Node) {
        if (!d.children || d3.active(this.svg.node())) return;
        this.isTransitioning = true;

        // Rebind the current node to the background.
        this.svg.select(".background").datum(d);

        // Define two sequenced transitions.
        const transition1 = this.svg.transition().duration(this.duration) as unknown as d3.Transition<d3.BaseType, unknown, null, undefined>;;
        const transition2 = transition1.transition();

        // Mark any currently-displayed bars as exiting.
        const exit = this.svg.selectAll(".enter")
            .attr("class", "exit");

        // Entering nodes immediately obscure the clicked-on bar, so hide it.
        exit.selectAll("rect")
            .attr("fill-opacity", p => p === d ? 0 : null);

        // Transition exiting bars to fade out.
        exit.transition(transition1)
            .attr("fill-opacity", 0)
            .remove();

        // Enter the new bars for the clicked-on data.
        // Per above, entering bars are immediately visible.
        const enter = this.bar(d, ".y-axis")
            .attr("fill-opacity", 0);

        // Have the text fade-in, even though the bars are visible.
        enter.transition(transition1)
            .attr("fill-opacity", 1);

        // Transition entering bars to their new y-position.
        enter.selectAll("g")
            .attr("transform", this.stack(d.index))
            .transition(transition1)
            .attr("transform", this.stagger());

        // Update the x-scale domain.
        this.x.domain([0, d3.max(d.children, d => d.value) ?? 0]);

        // Update the x-axis.
        this.svg.selectAll(".x-axis").transition(transition2)
            .call(this.xAxis);

        // Transition entering bars to the new x-scale.
        enter.selectAll("g").transition(transition2)
            .attr("transform", (d, i) => `translate(0,${this.barStep * i})`);

        // Color the bars as parents; they will fade to children if appropriate.
        enter.selectAll<SVGRectElement, Node>("rect")
            .attr("fill", this.color(true))
            .attr("fill-opacity", 1)
            .transition(transition2)
            .attr("fill", d => this.color(!!d.children))
            .attr("width", d => this.x(d.value ?? 0) - this.x(0));
        
        transition1.on("end", () => {
            this.isTransitioning = false;
        });
    }

    private stack(i: number) {
        let value = 0;
        return (d: any) => {
            const t = `translate(${this.x(value) - this.x(0)},${this.barStep * i})`;
            value += d.value;
            return t;
        };
    }

    private stagger() {
        let value = 0;
        return (d: any, i: number) => {
            const t = `translate(${this.x(value) - this.x(0)},${this.barStep * i})`;
            value += d.value;
            return t;
        };
    }

    private up(d: Node) {
        if (!d.parent || !this.svg.selectAll(".exit").empty()) return;
        this.isTransitioning = true;

        // Rebind the current node to the background.
        this.svg.select(".background").datum(d.parent);

        // Define two sequenced transitions.
        const transition1 = this.svg.transition().duration(this.duration) as unknown as d3.Transition<d3.BaseType, unknown, null, undefined>;;
        const transition2 = transition1.transition();

        // Mark any currently-displayed bars as exiting.
        const exit = this.svg.selectAll(".enter")
            .attr("class", "exit");

        // Update the x-scale domain.
        this.x.domain([0, d3.max((d.parent as Node).children as Node[], d => d.value) ?? 0]);

        // Update the x-axis.
        this.svg.selectAll(".x-axis").transition(transition1)
            .call(this.xAxis);

        // Transition exiting bars to the new x-scale.
        exit.selectAll("g").transition(transition1)
            .attr("transform", this.stagger());

        // Transition exiting bars to the parent’s position.
        exit.selectAll("g").transition(transition2)
            .attr("transform", this.stack(d.index));

        // Transition exiting rects to the new scale and fade to parent color.
        exit.selectAll<SVGRectElement, Node>("rect").transition(transition1)
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .attr("fill", this.color(true));

        // Transition exiting text to fade out.
        // Remove exiting nodes.
        exit.transition(transition2)
            .attr("fill-opacity", 0)
            .remove();

        // Enter the new bars for the clicked-on data's parent.
        const enter = this.bar(d.parent, ".exit")
            .attr("fill-opacity", 0);

        enter.selectAll("g")
            .attr("transform", (d, i) => `translate(0,${this.barStep * i})`);

        // Transition entering bars to fade in over the full duration.
        enter.transition(transition2)
            .attr("fill-opacity", 1);

        // Color the bars as appropriate.
        // Exiting nodes will obscure the parent bar, so hide it.
        // Transition entering rects to the new x-scale.
        // When the entering parent rect is done, make it visible!
        enter.selectAll<SVGRectElement, Node>("rect")
            .attr("fill", d => this.color(!!d.children))
            .attr("fill-opacity", p => p === d ? 0 : null)
            .transition(transition2)
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .on("end", function (p) { d3.select(this).attr("fill-opacity", 1); });
        
        transition1.on("end", () => {
            this.isTransitioning = false;
        });
    }
}