import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import visualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import * as d3 from "d3";
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
    private pendingUpdate: boolean = false;

    private getFilterTarget(level: number): { table: string, column: string } {
        const source = this.matrix!.rows.levels[level].sources[0];
        const queryName = source.queryName;
        const lastDot = queryName!.lastIndexOf('.');
        return {
            table: queryName!.substring(0, lastDot),
            column: queryName!.substring(lastDot + 1)
        };
    }

    private applyCurrentFilters() {
        this.pendingUpdate = true;

        if (this.levelFilters.size === 0) {
            this.host.applyJsonFilter(
                [], "general", "filter", powerbi.FilterAction.remove
            );
            return;
        }

        const filters = Array.from(this.levelFilters.entries())
            .sort(([a], [b]) => a - b) 
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
    
    private x!: d3.ScaleLinear<number, number>;
    private width = 0;
    private height = 0;

    private margin = { top: 30, right: 30, bottom: 0, left: 100 };

    private minBarStep = 10;
    private barStep = (d: Node) => {
        return d.children
            ? Math.max(
                Math.min(
                    (this.height - this.margin.top - this.margin.bottom) / d.children!.length, 
                    this.height / 3
                ), 
                this.minBarStep
            )
            : this.minBarStep
    }
    private barPadding = (barStep: number) => Math.max((2/30), 2 / barStep);

    private duration = 750;
    private isTransitioning = false;
    private color = d3.scaleOrdinal([true, false], ["steelblue", "#aaa"])

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

    private currentPath: string[] = [];

    private findNode(root: Node, path: string[]): Node {
        let current: Node = root;
        for (const name of path) {
            const match = current.children?.find(c => c.data.name === name) as Node | undefined;
            if (!match) return root;
            current = match;
        }
        return current;
    }

    public update(options: visualUpdateOptions) {
        // skip updates applied by filtering data due to selecting bars
        if (this.pendingUpdate) {
            this.pendingUpdate = false;
            return;
        }

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

        root.eachBefore(d => {
            (d as Node).index = d.parent
                ? d.parent.children!.indexOf(d)
                : 0;
        });

        this.width = options.viewport.width;
        this.height = options.viewport.height;

        this.x = d3.scaleLinear().range([this.margin.left, this.width - this.margin.right]);
        this.x.domain([0, root.value ?? 0]);

        this.svg.interrupt();
        this.svg.selectAll("*").remove();

        // a rebuild kills any in-flight transition, so never leave the flag stuck
        this.isTransitioning = false;

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

        const startNode = this.findNode(root, this.currentPath);
        this.down(startNode);
    }

    // layoutStep lets callers render the entering bars at the bar step that is
    // *currently on screen*, so the ratio transition can animate them to their
    // final step instead of them popping in at the final size.
    private bar(d: Node, selector: string, layoutStep?: number) {
        const barStep = layoutStep ?? this.barStep(d);
        const barPadding = this.barPadding(barStep)

        const g = this.svg.insert("g", selector)
            .attr("class", "enter")
            .attr("transform", `translate(0,${this.margin.top + barStep * barPadding})`)
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
            .attr("y", barStep * (1 - barPadding) / 2)
            .attr("dy", ".35em")
            .text(d => d.data.name);

        bar.append("rect")
            .attr("x", this.x(0))
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .attr("height", barStep * (1 - barPadding));
        return g;
    }

    private stack(i: number, barStep: number) {
        let value = 0;
        return (d: any) => {
            const t = `translate(${this.x(value) - this.x(0)},${barStep * i})`;
            value += d.value;
            return t;
        };
    }

    private stagger(barStep: number) {
        let value = 0;
        return (d: any, i: number) => {
            const t = `translate(${this.x(value) - this.x(0)},${barStep * i})`;
            value += d.value;
            return t;
        };
    }

    private ratio(barStep: number) {
        return (_d: any, i: number) => `translate(0,${barStep * i})`;
    }

    private down(d: Node) {
        if (!d.children || d3.active(this.svg.node())) return;

        // set a transition flag
        this.isTransitioning = true;

        // set the current new parent
        this.currentPath = d.ancestors().reverse().slice(1).map(n => n.data.name);

        // set barStep BEFORE transitions.
        // On the initial render d is the hierarchy root and has no parent,
        // so there is no "old" layout: fall back to the entering step.
        const ExitBarStep = d.parent ? this.barStep(d.parent as Node) : this.barStep(d);
        const ExitBarPadding = this.barPadding(ExitBarStep);
        const EnterBarStep = this.barStep(d);
        const EnterBarPadding = this.barPadding(EnterBarStep);

        // Rebind the current node to the background.
        this.svg.select(".background").datum(d);

        // If the bar step doesn't change, the ratio phase would tween every
        // attribute to the value it already has — skip it by zeroing its
        // duration. Chained transitions inherit timing, so the follow-up
        // phases must re-set their durations explicitly.
        const skipRatio = Math.abs(ExitBarStep - EnterBarStep) < 1e-6;

        // Define three sequenced transitions.
        const ratio_transition = this.svg.transition()
            .duration(skipRatio ? 0 : this.duration) as unknown as d3.Transition<d3.BaseType, unknown, null, undefined>;
        const stack_transition = ratio_transition.transition().duration(this.duration);
        const stagger_transition = stack_transition.transition();

        // Mark any currently-displayed bars as exiting.
        const exit = this.svg.selectAll(".enter")
            .attr("class", "exit");

        // Entering nodes immediately obscure the clicked-on bar, so hide it.
        exit.selectAll("rect")
            .attr("fill-opacity", p => p === d ? 0 : null)

        // adjust barstep of the exiting bars to the barstep of the bars entering
        exit.transition(ratio_transition)
                .attr("transform", `translate(0,${this.margin.top + EnterBarStep * EnterBarPadding})`);
            exit.selectAll("g").transition(ratio_transition)
                .attr("transform", this.ratio(EnterBarStep));
            exit.selectAll("rect").transition(ratio_transition)
                .attr("height", EnterBarStep * (1 - EnterBarPadding));
            exit.selectAll("text").transition(ratio_transition)
                .attr("y", EnterBarStep * (1 - EnterBarPadding) / 2);

        // Transition exiting bars to fade out.
        exit.transition(stack_transition)
            .attr("fill-opacity", 0)
            .remove();

        // Enter the new bars for the clicked-on data.
        // Per above, entering bars are immediately visible, so they must be
        // rendered at the OLD (exit) bar step to line up with the clicked bar.
        const enter = this.bar(d, ".y-axis", ExitBarStep)
            .attr("fill-opacity", 0);

        // Start stacked on the clicked bar in the old layout...
        enter.selectAll("g")
            .attr("transform", this.stack(d.index, ExitBarStep));

        // ...then resize and re-stack in lockstep with the exiting bars.
        enter.transition(ratio_transition)
            .attr("transform", `translate(0,${this.margin.top + EnterBarStep * EnterBarPadding})`);
        enter.selectAll("g").transition(ratio_transition)
            .attr("transform", this.stack(d.index, EnterBarStep));
        enter.selectAll("rect").transition(ratio_transition)
            .attr("height", EnterBarStep * (1 - EnterBarPadding));
        enter.selectAll("text").transition(ratio_transition)
            .attr("y", EnterBarStep * (1 - EnterBarPadding) / 2);

        // Have the text fade-in, even though the bars are visible.
        enter.transition(stack_transition)
            .attr("fill-opacity", 1);

        // Transition entering bars to their new y-position.
        enter.selectAll("g").transition(stack_transition)
            .attr("transform", this.stagger(EnterBarStep));

        // Update the x-scale domain.
        this.x.domain([0, d3.max(d.children, d => d.value) ?? 0]);

        // Update the x-axis.
        this.svg.selectAll(".x-axis").transition(stagger_transition)
            .call(this.xAxis);

        // Transition entering bars to the new x-scale.
        enter.selectAll("g").transition(stagger_transition)
            .attr("transform", (d, i) => `translate(0,${EnterBarStep * i})`);

        // Color the bars as parents; they will fade to children if appropriate.
        enter.selectAll<SVGRectElement, Node>("rect")
            .attr("fill", this.color(true))
            .attr("fill-opacity", 1)
            .transition(stagger_transition)
            .attr("fill", d => this.color(!!d.children))
            .attr("width", d => this.x(d.value ?? 0) - this.x(0));
        
        // reset transition flag (also on interrupt/cancel so clicks never lock up)
        stagger_transition.on("end interrupt cancel", () => {
            this.isTransitioning = false;
        });
    }

    private up(d: Node) {
        if (!d.parent || !this.svg.selectAll(".exit").empty()) return;
        this.currentPath = (d.parent as Node).ancestors().reverse().slice(1).map((n: any) => n.data.name);

        // set a transition flag
        this.isTransitioning = true;

        // set barStep BEFORE transitions
        const ExitBarStep = this.barStep(d);
        const ExitBarPadding = this.barPadding(ExitBarStep);
        const EnterBarStep = this.barStep(d.parent);
        const EnterBarPadding = this.barPadding(EnterBarStep);

         // Rebind the current node to the background.
        this.svg.select(".background").datum(d.parent);

        // If the bar step doesn't change, skip the (trailing) ratio phase.
        const skipRatio = Math.abs(ExitBarStep - EnterBarStep) < 1e-6;

        // Define two sequenced transitions.
        const stagger_transition = this.svg.transition().duration(this.duration) as unknown as d3.Transition<d3.BaseType, unknown, null, undefined>;;
        const stack_transition = stagger_transition.transition();
        const ratio_transition = stack_transition.transition()
            .duration(skipRatio ? 0 : this.duration);

        // Mark any currently-displayed bars as exiting.
        const exit = this.svg.selectAll(".enter")
            .attr("class", "exit");

        // Update the x-scale domain.
        this.x.domain([0, d3.max((d.parent as Node).children as Node[], d => d.value) ?? 0]);

        // Update the x-axis.
        this.svg.selectAll(".x-axis").transition(stagger_transition)
            .call(this.xAxis);

        // Transition exiting bars to the new x-scale.
        exit.selectAll("g").transition(stagger_transition)
            .attr("transform", this.stagger(ExitBarStep));

        // Transition exiting bars to the parent’s position.
        exit.selectAll("g").transition(stack_transition)
            .attr("transform", this.stack(d.index, ExitBarStep));

        // Transition exiting rects to the new scale and fade to parent color.
        exit.selectAll<SVGRectElement, Node>("rect").transition(stagger_transition)
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .attr("fill", this.color(true));

        // Transition exiting text to fade out.
        // Remove exiting nodes.
        exit.transition(stack_transition)
            .attr("fill-opacity", 0)
            .remove();

        // Enter the new bars for the clicked-on data's parent, rendered at the
        // OLD (exit) bar step; the ratio transition below resizes them.
        const enter = this.bar(d.parent, ".exit", ExitBarStep)
            .attr("fill-opacity", 0);

        enter.selectAll("g")
            .attr("transform", (_d, i) => `translate(0,${ExitBarStep * i})`);

        // Transition entering bars to fade in over the full duration.
        enter.transition(stack_transition)
            .attr("fill-opacity", 1);

        // Color the bars as appropriate.
        // Exiting nodes will obscure the parent bar, so hide it.
        // Transition entering rects to the new x-scale.
        // When the entering parent rect is done, make it visible!
        enter.selectAll<SVGRectElement, Node>("rect")
            .attr("fill", d => this.color(!!d.children))
            .attr("fill-opacity", p => p === d ? 0 : null)
            .transition(stack_transition)
            .attr("width", d => this.x(d.value ?? 0) - this.x(0))
            .on("end", function (p) { d3.select(this).attr("fill-opacity", 1); });
        
        enter.transition(ratio_transition)
            .attr("transform", `translate(0,${this.margin.top + EnterBarStep * EnterBarPadding})`);
        enter.selectAll("g").transition(ratio_transition)
            .attr("transform", this.ratio(EnterBarStep));
        enter.selectAll("rect").transition(ratio_transition)
            .attr("height", EnterBarStep * (1 - EnterBarPadding));
        enter.selectAll("text").transition(ratio_transition)
            .attr("y", EnterBarStep * (1 - EnterBarPadding) / 2);

        // reset transition flag (also on interrupt/cancel so clicks never lock up)
        ratio_transition.on("end interrupt cancel", () => {
            this.isTransitioning = false;
        });
    }
}