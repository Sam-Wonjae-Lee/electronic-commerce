//  Data Source: https://open.canada.ca/data/en/dataset/88078c92-d9e7-4123-8881-359fb3f5d608
//  Map image source reference: https://en.wikipedia.org/wiki/Provinces_and_territories_of_Canada
//  Have idea of the graph inside a screen 
//  Have a map of canada and highlight the provinces
//  Have the abstract map then allow people to hover to view info

const svgWidth = 1000;
const svgHeight = 500;

const DATA_SOURCE_URL = "https://open.canada.ca/data/en/dataset/88078c92-d9e7-4123-8881-359fb3f5d608";
const MAP_IMAGE_SOURCE_URL = "https://en.wikipedia.org/wiki/Provinces_and_territories_of_Canada";

let svg = d3.select("#map-area").append("svg")
            .attr("width", svgWidth)
            .attr("height", svgHeight);

let zoomToggleBtn = document.getElementById("zoom-toggle");
let zoomEnabled = true;

let tooltip = d3.select("body")
    .append("div")
    .style("position", "absolute")
    .style("background", "white")
    .style("padding", "10px")
    .style("border", "1px solid black")
    .style("border-radius", "5px")
    .style("pointer-events", "none")
    .style("opacity", 0);

const mapMaskConfig = {
    // Enabled: particles follow a bundled political map mask (not user-selected, not rendered).
    enabled: true,
    // Tries the first available file in this list.
    // (The repo currently includes `Political_map_of_Canada.svg.webp` rather than a raw `.svg`.)
    imageUrls: ["img/Political_map_of_Canada.svg", "img/Political_map_of_Canada.svg.webp"],
    tolerance: 26,
    assignTolerance: 135,
    downsample: 2,
    // How far we search from dark label/border pixels to find the underlying province fill colour.
    labelSearchRadius: 18
};

// Province masks (hand-drawn polygons in SVG coordinates).
// Supports single polygons (array of [x,y]) and multi-polygons (array of polygons).
// These are intentionally stylized but arranged to resemble Canada's provinces/regions.
let provinceShapes = {
    "British Columbia": [
        // Mainland (roughly matching a standard Canada provinces map layout)
        [70, 402], [55, 368], [55, 335], [62, 305], [75, 275], [95, 245],
        [125, 212], [155, 185], [182, 160], [210, 145], [236, 140],
        [255, 156], [265, 192], [262, 236], [250, 285], [236, 335],
        [210, 375], [170, 408], [120, 420]
    ],
    "Alberta": [
        [265, 192], [355, 185], [365, 392], [275, 404]
    ],
    "Manitoba and Saskatchewan": [
        [355, 185], [505, 185], [522, 402], [365, 404]
    ],
    "Ontario": [
        [505, 252], [535, 218], [590, 202], [650, 215], [695, 245],
        [715, 285], [705, 320], [675, 350], [632, 372], [585, 388],
        [540, 392], [512, 362], [495, 320]
    ],
    "Quebec": [
        [545, 210], [600, 150], [700, 122], [790, 150], [845, 212],
        [865, 268], [850, 320], [800, 362], [728, 392], [650, 398],
        [602, 382], [588, 338], [572, 300], [545, 262]
    ],
    "Atlantic provinces": [
        // NB + NS (combined stylized peninsula)
        [
            [835, 290], [875, 265], [915, 275], [930, 305], [925, 345],
            [900, 370], [865, 365], [845, 335]
        ],
        // PEI (island)
        [
            [900, 300], [915, 295], [925, 305], [910, 312]
        ],
        // Newfoundland and Labrador (stylized)
        [
            [910, 195], [950, 175], [985, 200], [985, 250], [955, 270],
            [920, 250]
        ]
    ]
};

let provinceColours = {
    "British Columbia": "#2f9e94",
    "Alberta": "#94a3b8",
    "Manitoba and Saskatchewan": "#86c56f",
    "Ontario": "#f4d03f",
    "Quebec": "#f2b6a0",
    "Atlantic provinces": "#d65a8a"
};

let totalOrdersData;
let averageOrdersData;
let averageValuePerPersonData;

d3.csv("data/22100073.csv").then(function(data){
    // Filter columns
    let filteredColumnData = data.map(function(d) {
        return {
            year: d["REF_DATE"],
            location: d["GEO"],
            order_destination: d["Destination of electronic orders"],
            order_type: d["Number and value of electronic orders"],
            unit_of_measure: d["UOM"],
            value: d["VALUE"]
        }
    })
    // console.log(filteredColumnData);
    
    // Convert csv values to numbers
    filteredColumnData.forEach(function(d) {
        d.value = +d.value
    })

    // Filter for electronic orders for companies in canada
    filteredColumnData = filteredColumnData.filter(function(d) {
        return d.order_destination == "Electronic orders to companies in Canada";
    })

    let provinceYearData = d3.group(
        filteredColumnData.filter(d => d.location !== "Canada"),
        d => d.year,
        d => d.location
    );

    function getProvinceData(year, provinceName) {
        let yearMap = provinceYearData.get(year);
        if (!yearMap) return {};

        let provinceRows = yearMap.get(provinceName) || [];
        let result = {};

        provinceRows.forEach(d => {
            if (d.order_type === "Number of orders") result.orders = d.value;
            if (d.order_type === "Average number of orders") result.avgOrders = d.value;
            if (d.order_type === "Value of orders") result.value = d.value;
            if (d.order_type === "Average value of orders per person") result.avgValue = d.value;
        });

        return result;
    }

    // Filter data for total number of orders
    totalOrdersData = filteredColumnData.filter(function(d) {
        return d.order_type == "Number of orders" && 
                d.unit_of_measure == "Number" && d.location != "Canada";
    })

    // Filter data for average number of orders
    averageOrdersData = filteredColumnData.filter(function(d) {
        return d.order_type == "Average number of orders" && 
                d.unit_of_measure == "Number" && d.location != "Canada";
    })

    // Filter data for average value of orders per person
    averageValuePerPersonData = filteredColumnData.filter(function(d) {
        return d.order_type == "Average value of orders per person" && 
                d.unit_of_measure == "Dollars" && d.location != "Canada";
    })

    let regionNames = Object.keys(provinceShapes);
    let viewModes = {
        map: "map",
        stacked: "stacked",
        stacked_area: "stacked_area",
        pie: "pie",
        bar: "bar"
    };

    let line = d3.line().curve(d3.curveLinearClosed);

    let chartLayer = svg.append("g").attr("class", "chart-layer");
    let viewport = svg.append("g").attr("class", "viewport");
    let bgLayer = viewport.append("g").attr("class", "bg-layer");
    let provinceLayer = viewport.append("g").attr("class", "province-layer");
    let particleLayer = viewport.append("g").attr("class", "particle-layer");
    let labelLayer = viewport.append("g").attr("class", "label-layer");
    let uiLayer = viewport.append("g").attr("class", "ui-layer");

    // Visible on-visualization credits (not affected by zoom/pan).
    let creditLayer = svg.append("g").attr("class", "credit-layer");
    renderVisualizationCredits();

    function renderVisualizationCredits() {
        creditLayer.selectAll("*").remove();

        let textLeft = 12;
        let textTop = svgHeight - 34;

        let text = creditLayer.append("text")
            .attr("x", textLeft)
            .attr("y", textTop)
            .attr("dominant-baseline", "hanging");

        text.append("tspan")
            .text(`Sources — Data: ${DATA_SOURCE_URL}`)
            .attr("x", textLeft)
            .attr("dy", 0);

        text.append("tspan")
            .text(`Map image: ${MAP_IMAGE_SOURCE_URL}`)
            .attr("x", textLeft)
            .attr("dy", 14);

        // Background panel sized to text for readability.
        let bbox = text.node().getBBox();
        creditLayer.insert("rect", "text")
            .attr("x", bbox.x - 8)
            .attr("y", bbox.y - 6)
            .attr("width", bbox.width + 16)
            .attr("height", bbox.height + 12)
            .attr("rx", 10)
            .attr("ry", 10);
    }

    let zoomTransform = d3.zoomIdentity;
    let zoomBehavior = d3.zoom()
        .scaleExtent([1, 8])
        .translateExtent([[0, 0], [svgWidth, svgHeight]])
        .extent([[0, 0], [svgWidth, svgHeight]])
        .on("zoom", (event) => {
            zoomTransform = event.transform;
            viewport.attr("transform", zoomTransform);
        });

    let zoomAttached = false;
    function enableZoom() {
        if (zoomAttached) return;
        svg.call(zoomBehavior);
        // Sync zoom's internal state with our current transform.
        svg.call(zoomBehavior.transform, zoomTransform);
        zoomAttached = true;
    }

    function disableZoom() {
        if (zoomAttached) svg.on(".zoom", null);
        zoomAttached = false;
    }

    function disableZoomAndReset() {
        if (zoomAttached) svg.on(".zoom", null);
        zoomAttached = false;
        zoomTransform = d3.zoomIdentity;
        viewport.attr("transform", null);
    }

    let mapMask = {
        ready: false,
        pointsByRegion: new Map(),
        imageFit: null
    };

    function polygonsForShape(shape) {
        if (!shape) return [];
        // Single polygon: [[x,y], ...]
        if (typeof shape[0]?.[0] === "number") return [shape];
        // Multi-polygon: [ [[x,y], ...], ... ]
        return shape;
    }

    function encodePointsToB64(points) {
        // uint16 little-endian pairs (x,y)
        let bytes = new Uint8Array(points.length * 4);
        for (let i = 0; i < points.length; i++) {
            let x = Math.max(0, Math.min(65535, Math.round(points[i][0]))) | 0;
            let y = Math.max(0, Math.min(65535, Math.round(points[i][1]))) | 0;
            let o = i * 4;
            bytes[o] = x & 0xff;
            bytes[o + 1] = (x >>> 8) & 0xff;
            bytes[o + 2] = y & 0xff;
            bytes[o + 3] = (y >>> 8) & 0xff;
        }
        let bin = "";
        for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
        return btoa(bin);
    }

    function decodePointsFromB64(b64) {
        let bin = atob(b64);
        let bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i) & 0xff;

        let pts = new Array(Math.floor(bytes.length / 4));
        for (let i = 0, p = 0; i + 3 < bytes.length; i += 4, p++) {
            let x = bytes[i] | (bytes[i + 1] << 8);
            let y = bytes[i + 2] | (bytes[i + 3] << 8);
            pts[p] = [x, y];
        }
        return pts;
    }

    function boundsForPolygon(polygon) {
        let xs = polygon.map(p => p[0]);
        let ys = polygon.map(p => p[1]);
        return {
            minX: Math.min(...xs),
            maxX: Math.max(...xs),
            minY: Math.min(...ys),
            maxY: Math.max(...ys)
        };
    }

    function randomPointInPolygon(polygon, maxAttempts = 2000) {
        let b = boundsForPolygon(polygon);
        for (let i = 0; i < maxAttempts; i++) {
            let x = Math.random() * (b.maxX - b.minX) + b.minX;
            let y = Math.random() * (b.maxY - b.minY) + b.minY;
            if (d3.polygonContains(polygon, [x, y])) return [x, y];
        }
        return null;
    }

    function randomPointInShape(shape) {
        let polygons = polygonsForShape(shape);
        if (!polygons.length) return null;

        let areas = polygons.map(p => Math.abs(d3.polygonArea(p)));
        let totalArea = d3.sum(areas);
        if (!Number.isFinite(totalArea) || totalArea <= 0) return randomPointInPolygon(polygons[0]);

        let pick = Math.random() * totalArea;
        let idx = 0;
        while (idx < areas.length - 1 && pick > areas[idx]) {
            pick -= areas[idx];
            idx += 1;
        }
        return randomPointInPolygon(polygons[idx]);
    }

    function centroidForShape(shape) {
        let polygons = polygonsForShape(shape);
        if (!polygons.length) return [0, 0];

        let weighted = polygons.map(p => {
            let area = Math.abs(d3.polygonArea(p));
            let c = d3.polygonCentroid(p);
            return { area, c };
        });
        let total = d3.sum(weighted, d => d.area);
        if (!Number.isFinite(total) || total <= 0) return d3.polygonCentroid(polygons[0]);

        let x = d3.sum(weighted, d => d.c[0] * d.area) / total;
        let y = d3.sum(weighted, d => d.c[1] * d.area) / total;
        return [x, y];
    }

    function pathForShape(shape) {
        let polygons = polygonsForShape(shape);
        return polygons.map(p => line(p)).join(" ");
    }

    function pointsByRegionFromB64Object(obj) {
        let m = new Map();
        regionNames.forEach(region => {
            let b64 = obj?.[region];
            m.set(region, b64 ? decodePointsFromB64(b64) : []);
        });
        return m;
    }

    function computeImageFit(imgWidth, imgHeight) {
        let scale = Math.min(svgWidth / imgWidth, svgHeight / imgHeight);
        let drawWidth = imgWidth * scale;
        let drawHeight = imgHeight * scale;
        let x = (svgWidth - drawWidth) / 2;
        let y = (svgHeight - drawHeight) / 2;
        return { x, y, width: drawWidth, height: drawHeight, scale };
    }

    function initMapMask() {
        if (!mapMaskConfig.enabled) return Promise.resolve(false);

        // Bump this when the mask-building algorithm changes, so old cached speckles disappear.
        let cacheKey = "canadaParticleMaskPoints:v8";
        try {
            let cached = localStorage.getItem(cacheKey);
            if (cached) {
                let parsed = JSON.parse(cached);
                mapMask.pointsByRegion = pointsByRegionFromB64Object(parsed);
                let totalPoints = d3.sum(Array.from(mapMask.pointsByRegion.values()), pts => pts.length);
                mapMask.ready = totalPoints > 2000;
                if (mapMask.ready) {
                    return Promise.resolve(true);
                }
            }
        } catch (e) {
            // ignore cache parse errors
        }

        function loadImage(url) {
            return new Promise((resolve, reject) => {
                let img = new Image();
                if (/^https?:\/\//i.test(url)) img.crossOrigin = "anonymous";
                img.onload = () => resolve(img);
                img.onerror = () => reject(new Error(`Failed to load ${url}`));
                img.src = url;
            });
        }

        async function loadFirstAvailable(urls) {
            for (let i = 0; i < urls.length; i++) {
                try {
                    let img = await loadImage(urls[i]);
                    return { img, url: urls[i] };
                } catch (e) {
                    // try next
                }
            }
            return null;
        }

        function colourDistance(a, b) {
            return Math.abs(a.r - b.r) + Math.abs(a.g - b.g) + Math.abs(a.b - b.b);
        }

        function isNearWhite(r, g, b) {
            return (r + g + b) >= 740;
        }

        function isNearBlack(r, g, b) {
            return (r + g + b) <= 28;
        }

        function isDarkInk(r, g, b) {
            return (r + g + b) <= 140;
        }

        function looksLikeInk(c) {
            if (c.a < 200) return false;
            // Many political SVG maps have black/gray borders and labels on white background.
            // Treat dark pixels as ink and try to pick a nearby province fill colour instead.
            return isDarkInk(c.r, c.g, c.b);
        }

        return (async () => {
            let loaded = await loadFirstAvailable(mapMaskConfig.imageUrls);
            if (!loaded) {
                mapMask.ready = false;
                return false;
            }

            let { img, url } = loaded;
            mapMask.imageUrl = url;

            let fit = computeImageFit(img.naturalWidth || img.width, img.naturalHeight || img.height);
            mapMask.imageFit = fit;

            let canvas = document.createElement("canvas");
            canvas.width = svgWidth;
            canvas.height = svgHeight;
            let ctx = canvas.getContext("2d", { willReadFrequently: true });
            ctx.clearRect(0, 0, svgWidth, svgHeight);
            ctx.drawImage(img, fit.x, fit.y, fit.width, fit.height);

            let imgData;
            try {
                imgData = ctx.getImageData(0, 0, svgWidth, svgHeight);
            } catch (e) {
                mapMask.ready = false;
                return false;
            }

            let pixels = imgData.data;
            let w = svgWidth;
            let h = svgHeight;

            function idxFor(x, y) {
                return (y * w + x) * 4;
            }

            function colorAt(x, y) {
                let i = idxFor(x, y);
                return { r: pixels[i], g: pixels[i + 1], b: pixels[i + 2], a: pixels[i + 3] };
            }

            // Detect whether the map has a black background (like the provided province map).
            let sampleCount = 0;
            let blackCount = 0;
            for (let y = 0; y < h; y += 20) {
                for (let x = 0; x < w; x += 20) {
                    let c = colorAt(x, y);
                    if (c.a < 200) continue;
                    sampleCount += 1;
                    if (isNearBlack(c.r, c.g, c.b)) blackCount += 1;
                }
            }
            let blackBackgroundMode = sampleCount > 0 && (blackCount / sampleCount) > 0.45;

            function isBackground(c) {
                if (c.a < 200) return true;
                if (isNearWhite(c.r, c.g, c.b)) return true;
                if (blackBackgroundMode && isNearBlack(c.r, c.g, c.b)) return true;
                return false;
            }

            function findRepresentativeFillColour(x0, y0, maxRadius) {
                for (let r = 1; r <= maxRadius; r++) {
                    for (let dy = -r; dy <= r; dy++) {
                        for (let dx = -r; dx <= r; dx++) {
                            if (Math.abs(dx) !== r && Math.abs(dy) !== r) continue;
                            let x = x0 + dx;
                            let y = y0 + dy;
                            if (x < 0 || x >= w || y < 0 || y >= h) continue;
                            let c = colorAt(x, y);
                            if (isBackground(c)) continue;
                            // Prefer brighter province fills over border/label ink.
                            if (blackBackgroundMode && isDarkInk(c.r, c.g, c.b)) continue;
                            return c;
                        }
                    }
                }
                return null;
            }

            function representativeColourForPixel(x, y) {
                let c = colorAt(x, y);
                if (isBackground(c)) return null;
                if (looksLikeInk(c)) {
                    return findRepresentativeFillColour(x, y, mapMaskConfig.labelSearchRadius);
                }
                return c;
            }

            function findSeedColour(seedX, seedY) {
                let maxRadius = 40;
                for (let r = 0; r <= maxRadius; r++) {
                    for (let dy = -r; dy <= r; dy++) {
                        for (let dx = -r; dx <= r; dx++) {
                            if (Math.abs(dx) !== r && Math.abs(dy) !== r) continue;
                            let x = Math.round(seedX + dx);
                            let y = Math.round(seedY + dy);
                            if (x < 0 || x >= w || y < 0 || y >= h) continue;
                            let rep = representativeColourForPixel(x, y);
                            if (!rep) continue;
                            return rep;
                        }
                    }
                }
                return null;
            }

            // Seeds are *relative* coordinates inside the fitted image rect.
            // Multiple seeds per region allow one dataset region to cover multiple province colours.
            let seedRatiosByRegion = {
                "British Columbia": [[0.135, 0.65]],
                "Alberta": [[0.25, 0.65]],
                "Manitoba and Saskatchewan": [[0.345, 0.65], [0.41, 0.65]],
                "Ontario": [[0.52, 0.744]],
                "Quebec": [[0.665, 0.624]],
                "Atlantic provinces": [[0.885, 0.62], [0.9, 0.72], [0.925, 0.68]],
                "__ignore__": [[0.14, 0.36], [0.315, 0.35], [0.47, 0.31], [0.54, 0.24]]
            };

            function seedFromRatio([rx, ry]) {
                return [
                    Math.round(fit.x + rx * fit.width),
                    Math.round(fit.y + ry * fit.height)
                ];
            }

            let seedsByRegion = {};
            Object.keys(seedRatiosByRegion).forEach(region => {
                seedsByRegion[region] = seedRatiosByRegion[region].map(seedFromRatio);
            });

            function dedupeColours(colours) {
                let unique = [];
                colours.forEach(c => {
                    if (!c) return;
                    if (!unique.some(u => colourDistance(u, c) < 18)) unique.push(c);
                });
                return unique;
            }

            let coloursByRegion = new Map();
            Object.keys(seedsByRegion).forEach(region => {
                let colours = seedsByRegion[region].map(([sx, sy]) => findSeedColour(sx, sy));
                coloursByRegion.set(region, dedupeColours(colours));
            });

            mapMask.pointsByRegion = new Map();
            regionNames.forEach(region => mapMask.pointsByRegion.set(region, []));

            let x0 = Math.max(0, Math.floor(fit.x));
            let y0 = Math.max(0, Math.floor(fit.y));
            let x1 = Math.min(w, Math.ceil(fit.x + fit.width));
            let y1 = Math.min(h, Math.ceil(fit.y + fit.height));

            let regions = regionNames.slice();
            if ((coloursByRegion.get("__ignore__") || []).length) regions.push("__ignore__");

            let ds = Math.max(1, mapMaskConfig.downsample);
            let assignTol = mapMaskConfig.assignTolerance;

            // Build a low-res label grid first so we can apply a simple majority filter to remove
            // speckle misclassifications (isolated wrong-colour pixels around borders/labels).
            let gridW = Math.max(1, Math.ceil((x1 - x0) / ds));
            let gridH = Math.max(1, Math.ceil((y1 - y0) / ds));
            let labels = new Int16Array(gridW * gridH);
            labels.fill(-1);

            let regionIndexByName = new Map();
            regionNames.forEach((r, i) => regionIndexByName.set(r, i));

            for (let y = y0; y < y1; y += ds) {
                for (let x = x0; x < x1; x += ds) {
                    let rep = representativeColourForPixel(x, y);
                    if (!rep) continue;

                    let bestRegion = null;
                    let bestDist = Infinity;

                    for (let i = 0; i < regions.length; i++) {
                        let region = regions[i];
                        let seedColours = coloursByRegion.get(region) || [];
                        if (!seedColours.length) continue;

                        let dMin = Infinity;
                        for (let j = 0; j < seedColours.length; j++) {
                            let d = colourDistance(rep, seedColours[j]);
                            if (d < dMin) dMin = d;
                        }

                        if (dMin < bestDist) {
                            bestDist = dMin;
                            bestRegion = region;
                        }
                    }

                    if (!bestRegion || bestDist > assignTol) continue;
                    if (bestRegion === "__ignore__") continue;

                    let gx = Math.floor((x - x0) / ds);
                    let gy = Math.floor((y - y0) / ds);
                    if (gx < 0 || gx >= gridW || gy < 0 || gy >= gridH) continue;
                    let idx = gy * gridW + gx;
                    labels[idx] = regionIndexByName.get(bestRegion) ?? -1;
                }
            }

            // Majority filter pass (3x3 neighborhood).
            let filtered = new Int16Array(labels.length);
            filtered.set(labels);
            let counts = new Int16Array(regionNames.length);

            for (let gy = 0; gy < gridH; gy++) {
                for (let gx = 0; gx < gridW; gx++) {
                    let idx = gy * gridW + gx;
                    let cur = labels[idx];
                    if (cur < 0) continue;

                    counts.fill(0);
                    for (let dy = -1; dy <= 1; dy++) {
                        let ny = gy + dy;
                        if (ny < 0 || ny >= gridH) continue;
                        for (let dx = -1; dx <= 1; dx++) {
                            let nx = gx + dx;
                            if (nx < 0 || nx >= gridW) continue;
                            let n = labels[ny * gridW + nx];
                            if (n >= 0) counts[n] += 1;
                        }
                    }

                    let best = cur;
                    let bestCount = counts[cur] || 0;
                    for (let i = 0; i < counts.length; i++) {
                        if (counts[i] > bestCount) {
                            best = i;
                            bestCount = counts[i];
                        }
                    }

                    // Only change when the neighborhood strongly agrees (prevents over-smoothing).
                    if (best !== cur && bestCount >= 5) filtered[idx] = best;
                }
            }

            // Remove tiny disconnected components per region (kills isolated outliers / speckle blobs).
            // Keep all components above a small absolute threshold and a small fraction of the region’s largest component.
            let cleaned = new Int16Array(filtered.length);
            cleaned.set(filtered);
            let visited = new Uint8Array(filtered.length);
            let stack = new Int32Array(filtered.length);

            function cleanRegionComponents(regionId) {
                visited.fill(0);
                let components = [];

                for (let gy = 0; gy < gridH; gy++) {
                    for (let gx = 0; gx < gridW; gx++) {
                        let startIdx = gy * gridW + gx;
                        if (visited[startIdx]) continue;
                        if (filtered[startIdx] !== regionId) continue;

                        // BFS (4-neighborhood) over grid cells.
                        let top = 0;
                        stack[top++] = startIdx;
                        visited[startIdx] = 1;

                        let cells = [];
                        while (top) {
                            let idx = stack[--top];
                            cells.push(idx);

                            let x = idx % gridW;
                            let y = (idx / gridW) | 0;

                            // 4-neighbors
                            if (x > 0) {
                                let n = idx - 1;
                                if (!visited[n] && filtered[n] === regionId) { visited[n] = 1; stack[top++] = n; }
                            }
                            if (x + 1 < gridW) {
                                let n = idx + 1;
                                if (!visited[n] && filtered[n] === regionId) { visited[n] = 1; stack[top++] = n; }
                            }
                            if (y > 0) {
                                let n = idx - gridW;
                                if (!visited[n] && filtered[n] === regionId) { visited[n] = 1; stack[top++] = n; }
                            }
                            if (y + 1 < gridH) {
                                let n = idx + gridW;
                                if (!visited[n] && filtered[n] === regionId) { visited[n] = 1; stack[top++] = n; }
                            }
                        }

                        components.push({ size: cells.length, cells });
                    }
                }

                if (!components.length) return;
                components.sort((a, b) => b.size - a.size);

                let largest = components[0].size;
                let minKeep = Math.max(12, Math.floor(largest * 0.0015)); // 0.15% of largest, but at least 12 cells.

                for (let i = 0; i < components.length; i++) {
                    if (components[i].size >= minKeep) continue;
                    let cells = components[i].cells;
                    for (let j = 0; j < cells.length; j++) cleaned[cells[j]] = -1;
                }

                // British Columbia can include a large island component (Vancouver Island),
                // but sometimes gets tiny misclassified blobs. Keep only the biggest components.
                let regionName = regionNames[regionId];
                if (regionName === "British Columbia" && components.length > 2) {
                    let keepAlways = 2; // mainland + Vancouver Island
                    let keepIfLarge = Math.max(minKeep, Math.floor(largest * 0.06)); // allow other substantial islands

                    for (let i = keepAlways; i < components.length; i++) {
                        if (components[i].size >= keepIfLarge) continue;
                        let cells = components[i].cells;
                        for (let j = 0; j < cells.length; j++) cleaned[cells[j]] = -1;
                    }
                }
            }

            for (let regionId = 0; regionId < regionNames.length; regionId++) {
                cleanRegionComponents(regionId);
            }

            // Fill small "holes" inside provinces that can happen when labels/borders are treated as background.
            // This assigns empty cells to the dominant surrounding region.
            let filled = new Int16Array(cleaned.length);
            filled.set(cleaned);

            function fillHolesOnce(src, dst) {
                dst.set(src);
                let localCounts = new Int16Array(regionNames.length);

                for (let gy = 0; gy < gridH; gy++) {
                    for (let gx = 0; gx < gridW; gx++) {
                        let idx = gy * gridW + gx;
                        if (src[idx] !== -1) continue;

                        localCounts.fill(0);
                        let considered = 0;
                        for (let dy = -3; dy <= 3; dy++) {
                            let ny = gy + dy;
                            if (ny < 0 || ny >= gridH) continue;
                            for (let dx = -3; dx <= 3; dx++) {
                                let nx = gx + dx;
                                if (nx < 0 || nx >= gridW) continue;
                                let v = src[ny * gridW + nx];
                                if (v < 0) continue;
                                considered += 1;
                                localCounts[v] += 1;
                            }
                        }

                        if (considered < 6) continue;

                        let best = -1;
                        let bestCount = 0;
                        for (let i = 0; i < localCounts.length; i++) {
                            if (localCounts[i] > bestCount) {
                                bestCount = localCounts[i];
                                best = i;
                            }
                        }

                        // Only fill when the neighborhood strongly agrees.
                        if (best >= 0 && bestCount >= 10 && bestCount / considered >= 0.7) {
                            dst[idx] = best;
                        }
                    }
                }
            }

            // Multiple passes fill larger label holes without bleeding across narrow borders.
            fillHolesOnce(filled, cleaned);
            fillHolesOnce(cleaned, filled);
            fillHolesOnce(filled, cleaned);
            fillHolesOnce(cleaned, filled);

            mapMask.pointsByRegion = new Map();
            regionNames.forEach(region => mapMask.pointsByRegion.set(region, []));

            for (let gy = 0; gy < gridH; gy++) {
                for (let gx = 0; gx < gridW; gx++) {
                    let idx = gy * gridW + gx;
                    let label = filled[idx];
                    if (label < 0) continue;
                    let region = regionNames[label];
                    mapMask.pointsByRegion.get(region)?.push([x0 + gx * ds, y0 + gy * ds]);
                }
            }

            // Prune isolated speckles that survived the majority pass (no same-region neighbors).
            regionNames.forEach(region => {
                let pts = mapMask.pointsByRegion.get(region) || [];
                if (pts.length < 40) return;

                let set = new Set(pts.map(p => `${p[0]},${p[1]}`));
                let kept = [];
                let offsets = [-ds, 0, ds];

                for (let i = 0; i < pts.length; i++) {
                    let x = pts[i][0];
                    let y = pts[i][1];
                    let hasNeighbor = false;

                    for (let dy = 0; dy < offsets.length && !hasNeighbor; dy++) {
                        for (let dx = 0; dx < offsets.length && !hasNeighbor; dx++) {
                            let ox = offsets[dx];
                            let oy = offsets[dy];
                            if (ox === 0 && oy === 0) continue;
                            if (set.has(`${x + ox},${y + oy}`)) hasNeighbor = true;
                        }
                    }

                    if (hasNeighbor) kept.push(pts[i]);
                }

                mapMask.pointsByRegion.set(region, kept);
            });

            // Region-specific cleanup: BC can sometimes pick up far-east speckles (e.g., above Quebec)
            // due to similar colours. Keep BC points to the west of Alberta's western edge (+margin).
            {
                let bcPts = mapMask.pointsByRegion.get("British Columbia") || [];
                let albertaShape = provinceShapes?.["Alberta"];

                if (bcPts.length && albertaShape) {
                    let albertaPoly = polygonsForShape(albertaShape)[0];
                    let b = boundsForPolygon(albertaPoly);
                    let maxBcX = b.minX + 25; // allow a small overlap into the AB border region

                    let kept = bcPts.filter(p => p[0] <= maxBcX);
                    // Only apply if we didn't accidentally delete most BC points.
                    if (kept.length >= 800 && kept.length / bcPts.length >= 0.6) {
                        mapMask.pointsByRegion.set("British Columbia", kept);
                    }
                }
            }

            regionNames.forEach(region => {
                let pts = mapMask.pointsByRegion.get(region) || [];
                if (pts.length) d3.shuffle(pts);
            });

            let totalPoints = d3.sum(Array.from(mapMask.pointsByRegion.values()), pts => pts.length);
            mapMask.ready = totalPoints > 2000;

            if (mapMask.ready) {
                try {
                    let payload = {};
                    regionNames.forEach(region => {
                        let pts = mapMask.pointsByRegion.get(region) || [];
                        payload[region] = pts.length ? encodePointsToB64(pts) : "";
                    });
                    localStorage.setItem(cacheKey, JSON.stringify(payload));
                } catch (e) {
                    // ignore storage quota/privacy errors
                }
            }
            return mapMask.ready;
        })();
    }

    function drawMapBackground() {
        bgLayer.selectAll("*").remove();
        // Intentionally blank: we use the map image only as an invisible mask.
    }

    function fnv1a32(str) {
        let hash = 0x811c9dc5;
        for (let i = 0; i < str.length; i++) {
            hash ^= str.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193) >>> 0;
        }
        return hash >>> 0;
    }

    function mulberry32(seed) {
        let t = seed >>> 0;
        return function() {
            t += 0x6d2b79f5;
            let x = Math.imul(t ^ (t >>> 15), 1 | t);
            x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
            return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
        };
    }

    function computeStackedAreaLayout(selectedMetric) {
        let cfg = metricConfig[selectedMetric] || metricConfig.total_orders;
        let years = ["2005", "2007", "2009"];

        let series = years.map(y => {
            let row = { year: +y };
            regionNames.forEach(region => {
                let info = getProvinceData(y, region);
                let v = cfg.getValue(info);
                row[region] = Number.isFinite(v) ? Math.max(0, v) : 0;
            });
            return row;
        });

        // Reserve a left gutter for the stacked-area legend so it doesn't overlap the plot.
        let baseLeft = 56;
        let padRight = 16;
        let minInnerW = 420;
        let legendReserve = Math.min(340, Math.max(0, svgWidth - baseLeft - padRight - minInnerW));
        let pad = { left: baseLeft + legendReserve, right: padRight, top: 22, bottom: 34 };
        let innerW = svgWidth - pad.left - pad.right;
        let innerH = svgHeight - pad.top - pad.bottom;

        let x = d3.scaleLinear()
            .domain(d3.extent(series, d => d.year))
            .range([pad.left, pad.left + innerW]);

        let stack = d3.stack().keys(regionNames);
        let layers = stack(series);
        let yMax = d3.max(layers, layer => d3.max(layer, d => d[1])) || 1;

        let y = d3.scaleLinear()
            .domain([0, yMax])
            .nice()
            .range([pad.top + innerH, pad.top]);

        return { cfg, years, series, pad, innerW, innerH, x, y, layers };
    }

    function drawStackedAreaChart(selectedYear, selectedMetric) {
        chartLayer.selectAll("*").remove();

        let { cfg, pad, innerW, innerH, x, y, layers, years } = computeStackedAreaLayout(selectedMetric);

        let area = d3.area()
            .x(d => x(d.data.year))
            .y0(d => y(d[0]))
            .y1(d => y(d[1]))
            .curve(d3.curveMonotoneX);

        // Background and frame
        chartLayer.append("rect")
            .attr("x", pad.left)
            .attr("y", pad.top)
            .attr("width", innerW)
            .attr("height", innerH)
            .attr("fill", "#ffffff")
            .attr("stroke", "#e5e7eb");

        chartLayer.selectAll("path.area")
            .data(layers, d => d.key)
            .enter()
            .append("path")
            .attr("class", "area")
            .attr("d", d => area(d))
            .attr("fill", d => provinceColours[d.key] || "#999")
            .attr("fill-opacity", 0.18)
            .attr("stroke", "rgba(255,255,255,0.85)")
            .attr("stroke-width", 0.6);

        // Axes (simple)
        let xAxis = d3.axisBottom(x).ticks(3).tickFormat(d3.format("d"));
        let yFmt = metricFormat[selectedMetric]?.axis || d3.format(".2s");
        let yAxis = d3.axisLeft(y).ticks(5).tickFormat(yFmt);

        chartLayer.append("g")
            .attr("transform", `translate(0,${pad.top + innerH})`)
            .call(xAxis)
            .call(g => g.selectAll("text").attr("font-family", "system-ui").attr("font-size", 11))
            .call(g => g.selectAll("path,line").attr("stroke", "#9ca3af"));

        chartLayer.append("g")
            .attr("transform", `translate(${pad.left},0)`)
            .call(yAxis)
            .call(g => g.selectAll("text").attr("font-family", "system-ui").attr("font-size", 11))
            .call(g => g.selectAll("path,line").attr("stroke", "#9ca3af"));

        chartLayer.append("text")
            .attr("x", pad.left)
            .attr("y", 14)
            .attr("fill", "#111827")
            .attr("font-family", "system-ui")
            .attr("font-size", 12)
            .attr("font-weight", 600)
            .text((() => {
                if (selectedMetric === "avg_orders") return "Stacked area: Average number of orders";
                if (selectedMetric === "value_orders") return "Stacked area: Value of orders";
                if (selectedMetric === "avg_value_per_person") return "Stacked area: Average value of orders per person";
                return "Stacked area: Total number of orders";
            })());

        // Legend (province colours + selected-year share)
        (function drawStackedAreaLegend() {
            let yearStr = String(selectedYear || "");
            let total = 0;
            let entries = regionNames.map(region => {
                let info = getProvinceData(yearStr, region);
                let v = cfg.getValue(info);
                v = Number.isFinite(v) ? Math.max(0, v) : 0;
                total += v;
                return { region, value: v };
            });
            if (!(total > 0)) return;

            entries = entries
                .filter(d => d.value > 0)
                .sort((a, b) => b.value - a.value)
                .map(d => ({ ...d, pct: (d.value / total) * 100 }));
            if (!entries.length) return;

            let padL = 10;
            let rowH = 18;
            let titleH = 18;
            let legendW = Math.min(320, Math.max(160, pad.left - 24));
            let legendH = padL * 2 + titleH + entries.length * rowH;

            // Place legend on the left side of the chart.
            let x0 = 12;
            let y0 = pad.top + 6;
            if (y0 + legendH > svgHeight - 10) y0 = Math.max(10, svgHeight - legendH - 10);

            let g = chartLayer.append("g").attr("class", "stacked-area-legend");
            g.append("rect")
                .attr("x", x0)
                .attr("y", y0)
                .attr("width", legendW)
                .attr("height", legendH)
                .attr("rx", 10)
                .attr("fill", "rgba(255,255,255,0.92)");

            let fmt = d3.format(".1f");
            g.append("text")
                .attr("x", x0 + padL)
                .attr("y", y0 + padL + 12)
                .attr("fill", "#111827")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .attr("font-weight", 650)
                .text(`Province share (%), ${yearStr}`);

            entries.forEach((d, i) => {
                let yy = y0 + padL + titleH + i * rowH + 12;
                g.append("rect")
                    .attr("x", x0 + padL)
                    .attr("y", yy - 10)
                    .attr("width", 10)
                    .attr("height", 10)
                    .attr("rx", 2)
                    .attr("fill", provinceColours[d.region] || "#6b7280");

                g.append("text")
                    .attr("x", x0 + padL + 16)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .text(d.region);

                g.append("text")
                    .attr("x", x0 + legendW - padL)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .attr("text-anchor", "end")
                    .text(`${fmt(d.pct)}%`);
            });
        })();

        // Selected-year marker
        let yearNum = +selectedYear;
        if (Number.isFinite(yearNum)) {
            let mx = x(yearNum);
            chartLayer.append("line")
                .attr("x1", mx)
                .attr("x2", mx)
                .attr("y1", pad.top)
                .attr("y2", pad.top + innerH)
                .attr("stroke", "#111827")
                .attr("stroke-width", 1)
                .attr("opacity", 0.75);
        }

        // Click/drag to choose year (snaps to the 3 years).
        let yearsNum = years.map(d => +d);
        chartLayer.append("rect")
            .attr("x", pad.left)
            .attr("y", pad.top)
            .attr("width", innerW)
            .attr("height", innerH)
            .attr("fill", "transparent")
            .style("cursor", "pointer")
            .on("click", function(event) {
                let [mx] = d3.pointer(event, svg.node());
                let yr = x.invert(mx);
                let nearest = yearsNum.reduce((best, cur) => (Math.abs(cur - yr) < Math.abs(best - yr) ? cur : best), yearsNum[0]);
                if (typeof setYearValue === "function") setYearValue(String(nearest));
            });
    }

    function drawProvinces(viewMode) {
        provinceLayer.selectAll("*").remove();

        if (viewMode !== viewModes.map) return;

        if (mapMask.ready) {
            drawMapBackground();
            return;
        }

        provinceLayer.selectAll(".province")
            .data(regionNames, d => d)
            .enter()
            .append("path")
            .attr("class", "province")
            .attr("d", d => pathForShape(provinceShapes[d]))
            .attr("fill", d => provinceColours[d] || "#ccc")
            .attr("fill-opacity", 0.08)
            .attr("stroke", "#333")
            .attr("stroke-width", 1);
    }

    function drawLabels(viewMode) {
        labelLayer.selectAll("*").remove();

        if (viewMode !== viewModes.map) return;
        if (mapMask.ready) return;

        labelLayer.selectAll(".province-label")
            .data(regionNames, d => d)
            .enter()
            .append("text")
            .attr("class", "province-label")
            .attr("x", d => centroidForShape(provinceShapes[d])[0])
            .attr("y", d => centroidForShape(provinceShapes[d])[1])
            .text(d => d)
            .attr("fill", "#111")
            .attr("text-anchor", "middle")
            .attr("dominant-baseline", "middle")
            .attr("font-size", "12px");
    }

    function drawMapLegend() {
        let entries = regionNames.map(region => ({
            region,
            colour: provinceColours[region] || "#6b7280"
        }));

        let pad = 10;
        let rowH = 18;
        let titleH = 18;
        let legendW = 260;
        let legendH = pad * 2 + titleH + entries.length * rowH;

        let x = 12;
        let y = Math.max(10, Math.min(12, svgHeight - legendH - 10));

        let g = chartLayer.append("g").attr("class", "map-legend");
        g.append("rect")
            .attr("x", x)
            .attr("y", y)
            .attr("width", legendW)
            .attr("height", legendH)
            .attr("rx", 10)
            .attr("fill", "rgba(255,255,255,0.92)");

        g.append("text")
            .attr("x", x + pad)
            .attr("y", y + pad + 12)
            .attr("fill", "#111827")
            .attr("font-family", "system-ui")
            .attr("font-size", 12)
            .attr("font-weight", 650)
            .text("Provinces");

        entries.forEach((d, i) => {
            let yy = y + pad + titleH + i * rowH + 12;
            g.append("rect")
                .attr("x", x + pad)
                .attr("y", yy - 10)
                .attr("width", 10)
                .attr("height", 10)
                .attr("rx", 2)
                .attr("fill", d.colour);

            g.append("text")
                .attr("x", x + pad + 16)
                .attr("y", yy)
                .attr("fill", "#111827")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .text(d.region);
        });
    }


    let metricConfig = {
        total_orders: {
            getValue: (info) => info.orders,
            // Particles are: base fill (from province area) + extras (from metric values).
            extraParticles: 11000,
            weightTransform: (v) => v
        },
        avg_orders: {
            getValue: (info) => info.avgOrders,
            extraParticles: 11000,
            weightTransform: (v) => v
        },
        value_orders: {
            getValue: (info) => info.value,
            extraParticles: 11000,
            weightTransform: (v) => v
        },
        avg_value_per_person: {
            getValue: (info) => info.avgValue,
            extraParticles: 11000,
            weightTransform: (v) => v
        }
    };

    let metricFormat = {
        total_orders: {
            axis: d3.format(",.0f"),
            value: d3.format(",.0f")
        },
        avg_orders: {
            axis: d3.format(",.1f"),
            value: d3.format(",.1f")
        },
        value_orders: {
            axis: d3.format("$,.0f"),
            value: d3.format("$,.0f")
        },
        avg_value_per_person: {
            axis: d3.format("$,.0f"),
            value: d3.format("$,.1f")
        }
    };

    function formatMetricValue(metricKey, v) {
        if (!Number.isFinite(v)) return "N/A";
        let fmt = metricFormat[metricKey]?.value;
        if (!fmt) return String(v);
        return fmt(v);
    }

    function metricTitle(metricKey) {
        if (metricKey === "avg_orders") return "Average number of orders";
        if (metricKey === "value_orders") return "Value of orders";
        if (metricKey === "avg_value_per_person") return "Average value of orders per person";
        return "Total number of orders";
    }

    function getRegionMetricEntriesForYear(yearStr, metricKey) {
        let cfg = metricConfig[metricKey] || metricConfig.total_orders;
        return regionNames.map(region => {
            let info = getProvinceData(String(yearStr), region);
            let v = cfg.getValue(info);
            v = Number.isFinite(v) ? Math.max(0, v) : 0;
            return { region, value: v };
        });
    }

    function showChartTooltip(event, html) {
        tooltip
            .style("opacity", 1)
            .html(html)
            .style("left", (event.pageX + 10) + "px")
            .style("top", (event.pageY + 10) + "px");
    }

    function hideChartTooltip() {
        tooltip.style("opacity", 0);
    }

    function updateFactsPanel(selectedYear, selectedMetric) {
        let panel = document.getElementById("facts-panel");
        if (!panel) return;

        let yearStr = String(selectedYear || "");
        let metricKey = selectedMetric || "total_orders";
        let title = metricTitle(metricKey);
        let fmt = metricFormat[metricKey]?.value || (v => String(v));

        let metricBlurb = {
            total_orders: "Counts how many electronic orders businesses made to companies in Canada (number of orders).",
            avg_orders: "Average number of electronic orders per business (number of orders).",
            value_orders: "Total dollar value of electronic orders (dollars).",
            avg_value_per_person: "Average dollar value of electronic orders per person (dollars)."
        };

        let entries = getRegionMetricEntriesForYear(yearStr, metricKey);
        let total = d3.sum(entries, d => d.value);

        let max = entries.reduce((best, cur) => (cur.value > best.value ? cur : best), entries[0] || { region: "", value: 0 });
        let nonZero = entries.filter(d => d.value > 0);
        let minNonZero = nonZero.reduce((best, cur) => (cur.value < best.value ? cur : best), nonZero[0] || null);

        let overviewFacts = [];
        let leadersFacts = [];
        let changeFacts = [];

        overviewFacts.push(`<li><strong>What it measures:</strong> ${metricBlurb[metricKey] || "Metric details unavailable."}</li>`);
        if (total > 0) {
            overviewFacts.push(`<li><strong>Total across provinces:</strong> ${fmt(total)}</li>`);

            leadersFacts.push(`<li><strong>Highest province:</strong> ${max.region} (${fmt(max.value)})</li>`);
            if (minNonZero) leadersFacts.push(`<li><strong>Lowest non-zero province:</strong> ${minNonZero.region} (${fmt(minNonZero.value)})</li>`);
            let share = (max.value / total) * 100;
            if (Number.isFinite(share)) leadersFacts.push(`<li><strong>Largest share:</strong> ${max.region} is ${d3.format(".1f")(share)}% of the total</li>`);
        } else {
            overviewFacts.push(`<li class="facts-muted">No non-zero values for this metric/year.</li>`);
        }

        let yearOrder = ["2005", "2007", "2009"];
        let idx = yearOrder.indexOf(yearStr);
        let prevYear = idx > 0 ? yearOrder[idx - 1] : null;
        if (prevYear) {
            let prevEntries = getRegionMetricEntriesForYear(prevYear, metricKey);
            let prevTotal = d3.sum(prevEntries, d => d.value);
            let deltaTotal = total - prevTotal;
            let pctTotal = (prevTotal > 0) ? (deltaTotal / prevTotal) * 100 : null;

            if (Number.isFinite(deltaTotal)) {
                changeFacts.push(`<li><strong>Total change vs ${prevYear}:</strong> ${fmt(deltaTotal)} (${pctTotal == null ? "N/A" : (d3.format("+.1f")(pctTotal) + "%")})</li>`);
            }

            let prevByRegion = new Map(prevEntries.map(d => [d.region, d.value]));
            let bestUp = null;
            let bestDown = null;
            entries.forEach(d => {
                let pv = prevByRegion.get(d.region) ?? 0;
                let delta = d.value - pv;
                if (bestUp == null || delta > bestUp.delta) bestUp = { region: d.region, delta, prev: pv, cur: d.value };
                if (bestDown == null || delta < bestDown.delta) bestDown = { region: d.region, delta, prev: pv, cur: d.value };
            });

            if (bestUp && Number.isFinite(bestUp.delta) && bestUp.delta > 0) {
                let pct = (bestUp.prev > 0) ? (bestUp.delta / bestUp.prev) * 100 : null;
                changeFacts.push(`<li><strong>Biggest increase:</strong> ${bestUp.region} (${fmt(bestUp.delta)}; ${pct == null ? "N/A" : (d3.format("+.1f")(pct) + "%")})</li>`);
            }
            if (bestDown && Number.isFinite(bestDown.delta) && bestDown.delta < 0) {
                let pct = (bestDown.prev > 0) ? (bestDown.delta / bestDown.prev) * 100 : null;
                changeFacts.push(`<li><strong>Biggest decrease:</strong> ${bestDown.region} (${fmt(bestDown.delta)}; ${pct == null ? "N/A" : (d3.format("+.1f")(pct) + "%")})</li>`);
            }
        } else {
            changeFacts.push(`<li class="facts-muted">No earlier year to compare (timeline starts at ${yearStr}).</li>`);
        }

        function card(title, items) {
            let list = items.length ? `<ul class="fact-card-list">${items.join("")}</ul>` : `<div class="facts-muted">No facts available.</div>`;
            return `<div class="fact-card"><div class="fact-card-title">${title}</div>${list}</div>`;
        }

        panel.innerHTML = `
            <div class="facts-header">
                <div class="facts-title">Interesting facts</div>
                <div class="facts-subtitle">${title} · ${yearStr}</div>
            </div>
            <div class="facts-grid">
                ${card("Overview", overviewFacts)}
                ${card("Leaders", leadersFacts)}
                ${card("Change", changeFacts)}
            </div>
        `;
    }

    function drawPieChart(selectedYear, selectedMetric) {
        let entries = getRegionMetricEntriesForYear(selectedYear, selectedMetric)
            .filter(d => d.value > 0)
            .sort((a, b) => b.value - a.value);

        let total = d3.sum(entries, d => d.value);
        if (!(total > 0) || !entries.length) {
            uiLayer.append("text")
                .attr("x", 12)
                .attr("y", 18)
                .attr("fill", "#111827")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .attr("font-weight", 650)
                .text(`Pie chart: ${metricTitle(selectedMetric)} (${selectedYear}) — no data`);
            return;
        }

        let legendW = 330;
        let cx = Math.max(legendW + 230, svgWidth * 0.62);
        let cy = svgHeight * 0.52;
        let outerR = Math.min(190, Math.min(svgWidth - cx - 18, svgHeight * 0.44));
        let innerR = outerR * 0.58;

        uiLayer.append("text")
            .attr("x", 12)
            .attr("y", 18)
            .attr("fill", "#111827")
            .attr("font-family", "system-ui")
            .attr("font-size", 12)
            .attr("font-weight", 650)
            .text(`Pie chart: ${metricTitle(selectedMetric)} (${selectedYear})`);

        let fmtPct = d3.format(".1f");

        // Center label (particles form the donut; this is just the label)
        uiLayer.append("text")
            .attr("text-anchor", "middle")
            .attr("x", cx)
            .attr("y", cy - 4)
            .attr("fill", "#111827")
            .attr("font-family", "system-ui")
            .attr("font-size", 12)
            .attr("font-weight", 700)
            .text(String(selectedYear));

        uiLayer.append("text")
            .attr("text-anchor", "middle")
            .attr("x", cx)
            .attr("y", cy + 14)
            .attr("fill", "#374151")
            .attr("font-family", "system-ui")
            .attr("font-size", 11)
            .text(metricTitle(selectedMetric));

        (function drawPieLegend() {
            let pad = 10;
            let rowH = 18;
            let titleH = 18;
            let legendX = 12;
            let legendY = 32;
            let legendH = pad * 2 + titleH + entries.length * rowH;

            let box = uiLayer.append("g").attr("class", "pie-legend");
            box.append("rect")
                .attr("x", legendX)
                .attr("y", legendY)
                .attr("width", legendW)
                .attr("height", legendH)
                .attr("rx", 10)
                .attr("fill", "rgba(255,255,255,0.92)");

            box.append("text")
                .attr("x", legendX + pad)
                .attr("y", legendY + pad + 12)
                .attr("fill", "#111827")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .attr("font-weight", 650)
                .text("Province share (%)");

            entries.forEach((d, i) => {
                let yy = legendY + pad + titleH + i * rowH + 12;
                let pct = (d.value / total) * 100;

                box.append("rect")
                    .attr("x", legendX + pad)
                    .attr("y", yy - 10)
                    .attr("width", 10)
                    .attr("height", 10)
                    .attr("rx", 2)
                    .attr("fill", provinceColours[d.region] || "#6b7280");

                box.append("text")
                    .attr("x", legendX + pad + 16)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .text(d.region);

                box.append("text")
                    .attr("x", legendX + legendW - pad)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .attr("text-anchor", "end")
                    .text(`${fmtPct(pct)}%`);
            });
        })();
    }

    function drawBarChart(selectedYear, selectedMetric) {
        let entries = getRegionMetricEntriesForYear(selectedYear, selectedMetric)
            .filter(d => d.value > 0)
            .sort((a, b) => b.value - a.value);

        let pad = { left: 210, right: 22, top: 34, bottom: 34 };
        let innerW = svgWidth - pad.left - pad.right;
        let innerH = svgHeight - pad.top - pad.bottom;

        uiLayer.append("text")
            .attr("x", 12)
            .attr("y", 18)
            .attr("fill", "#111827")
            .attr("font-family", "system-ui")
            .attr("font-size", 12)
            .attr("font-weight", 650)
            .text(`Bar chart: ${metricTitle(selectedMetric)} (${selectedYear})`);

        if (!entries.length) {
            uiLayer.append("text")
                .attr("x", 12)
                .attr("y", 40)
                .attr("fill", "#6b7280")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .text("No non-zero values for this year/metric.");
            return;
        }

        let maxV = d3.max(entries, d => d.value) || 1;
        let x = d3.scaleLinear().domain([0, maxV]).nice().range([pad.left, pad.left + innerW]);
        let y = d3.scaleBand().domain(entries.map(d => d.region)).range([pad.top, pad.top + innerH]).padding(0.22);

        bgLayer.append("rect")
            .attr("x", pad.left)
            .attr("y", pad.top)
            .attr("width", innerW)
            .attr("height", innerH)
            .attr("fill", "#ffffff")
            .attr("stroke", "#e5e7eb");

        let xAxis = d3.axisBottom(x).ticks(5).tickFormat(metricFormat[selectedMetric]?.axis || d3.format(".2s"));
        let yAxis = d3.axisLeft(y).tickSize(0);

        uiLayer.append("g")
            .attr("transform", `translate(0,${pad.top + innerH})`)
            .call(xAxis)
            .call(g => g.selectAll("text").attr("font-family", "system-ui").attr("font-size", 11))
            .call(g => g.selectAll("path,line").attr("stroke", "#9ca3af"));

        uiLayer.append("g")
            .attr("transform", `translate(${pad.left},0)`)
            .call(yAxis)
            .call(g => g.selectAll("text").attr("font-family", "system-ui").attr("font-size", 12))
            .call(g => g.selectAll("path").remove());
    }

    function drawParticles(selectedYear, selectedMetric, viewMode) {
        let cfg = metricConfig[selectedMetric] || metricConfig.total_orders;
        let particles = [];

        function relocateIsolatedParticles(particles) {
            // Fix stray single particles by moving them to a denser spot inside their own region mask.
            if (viewMode !== viewModes.map || !mapMask.ready) return;

            let cellSize = 8;
            let radius = 14;
            let radius2 = radius * radius;
            let minNeighbors = 2;

            function cellKey(cx, cy) {
                return `${cx},${cy}`;
            }

            function buildIndex() {
                let index = new Map();
                for (let i = 0; i < particles.length; i++) {
                    let p = particles[i];
                    let cx = Math.floor(p.x / cellSize);
                    let cy = Math.floor(p.y / cellSize);
                    let key = cellKey(cx, cy);
                    let arr = index.get(key);
                    if (!arr) index.set(key, (arr = []));
                    arr.push(i);
                }
                return index;
            }

            function neighborCountAt(index, x, y, region, ignoreIndex = -1) {
                let cx = Math.floor(x / cellSize);
                let cy = Math.floor(y / cellSize);
                let count = 0;

                for (let dy = -2; dy <= 2; dy++) {
                    for (let dx = -2; dx <= 2; dx++) {
                        let arr = index.get(cellKey(cx + dx, cy + dy));
                        if (!arr) continue;
                        for (let k = 0; k < arr.length; k++) {
                            let j = arr[k];
                            if (j === ignoreIndex) continue;
                            let pj = particles[j];
                            if (pj.location !== region) continue;
                            let ddx = pj.x - x;
                            let ddy = pj.y - y;
                            if (ddx * ddx + ddy * ddy <= radius2) count++;
                            if (count >= minNeighbors) return count;
                        }
                    }
                }
                return count;
            }

            let index = buildIndex();
            let outliers = [];

            for (let i = 0; i < particles.length; i++) {
                let p = particles[i];
                // Skip regions with too few particles; isolation is unavoidable there.
                // (Counts are already clamped up for masked regions, so this should rarely trigger.)
                if (!p.location) continue;
                let count = neighborCountAt(index, p.x, p.y, p.location, i);
                if (count < minNeighbors) outliers.push(i);
            }

            if (!outliers.length) return;

            for (let oi = 0; oi < outliers.length; oi++) {
                let i = outliers[oi];
                let p = particles[i];
                let pts = mapMask.pointsByRegion.get(p.location) || [];
                if (pts.length < 10) continue;

                // Try a few candidate points until we land near other same-region particles.
                let best = null;
                let bestScore = -1;
                for (let t = 0; t < 25; t++) {
                    let cand = pts[(Math.random() * pts.length) | 0];
                    let score = neighborCountAt(index, cand[0], cand[1], p.location, i);
                    if (score > bestScore) {
                        bestScore = score;
                        best = cand;
                        if (score >= minNeighbors) break;
                    }
                }

                if (best) {
                    p.x = best[0];
                    p.y = best[1];
                }
            }
        }

        function computeCountsByRegion() {
            let rows = regionNames.map(region => {
                let info = getProvinceData(selectedYear, region);
                let raw = cfg.getValue(info);
                let v = Number.isFinite(raw) ? Math.max(0, raw) : 0;
                let w = cfg.weightTransform ? cfg.weightTransform(v) : v;
                if (!Number.isFinite(w) || w < 0) w = 0;
                return { region, v, w };
            });

            let totalW = d3.sum(rows, d => d.w);
            let extraBudget = cfg.extraParticles ?? cfg.totalParticles ?? 12000;
            let minPerRegion = cfg.minPerRegion ?? 120;
            let maxPerRegion = cfg.maxPerRegion ?? 9000;

            // Base fill: ensures provinces don't look "holey" even for small values.
            // Only applies to map view with an active mask.
            let baseByRegion = new Map();
            if (viewMode === viewModes.map && mapMask.ready) {
                rows.forEach(d => {
                    let pts = mapMask.pointsByRegion.get(d.region) || [];
                    // Smaller baseline so metric choice has a visible effect on particle counts.
                    let base = Math.floor(pts.length / 8);
                    base = Math.max(120, Math.min(base, 1800));
                    // If there’s literally no data, still keep a small silhouette.
                    if (d.v <= 0) base = Math.min(base, 400);
                    // Cap by available points.
                    if (pts.length) base = Math.min(base, pts.length);
                    baseByRegion.set(d.region, base);
                });
                // Extras should be smaller if we already have a large base fill.
                minPerRegion = Math.min(minPerRegion, 40);
            } else {
                rows.forEach(d => baseByRegion.set(d.region, 0));
            }

            // If a metric/year has no data, fall back to just the base fill (or small uniform if no mask).
            if (!(totalW > 0)) {
                let out = new Map();
                if (viewMode === viewModes.map && mapMask.ready) {
                    rows.forEach(d => out.set(d.region, baseByRegion.get(d.region) || 0));
                    return out;
                }
                let per = Math.max(0, Math.floor(extraBudget / Math.max(1, rows.length)));
                rows.forEach(d => out.set(d.region, per));
                return out;
            }

            let counts = rows.map(d => {
                let share = d.w / totalW;
                let extra = Math.round(share * extraBudget);

                if (d.v > 0) extra = Math.max(minPerRegion, extra);
                extra = Math.min(maxPerRegion, extra);
                let base = baseByRegion.get(d.region) || 0;
                let c = base + extra;
                return { region: d.region, count: c };
            });

            // Cap by available mask points (prevents oversampling a small province).
            if (mapMask.ready) {
                counts.forEach(d => {
                    let pts = mapMask.pointsByRegion.get(d.region) || [];
                    if (pts.length) d.count = Math.min(d.count, pts.length);
                });
            }

            // We don't force a strict global budget anymore once base fill is applied;
            // this avoids reintroducing holes via post-adjustment.

            let out = new Map();
            counts.forEach(d => out.set(d.region, d.count));
            return out;
        }

        function pointsForRegion(region, count) {
            if (!mapMask.ready) return null;
            let pts = mapMask.pointsByRegion.get(region) || [];
            if (!pts.length) return null;

            let out = [];
            let start = (Math.random() * pts.length) | 0;
            let step = Math.max(1, Math.floor(pts.length / count));
            for (let i = 0; i < count; i++) {
                out.push(pts[(start + i * step) % pts.length]);
            }
            return out;
        }

        let countsByRegion = computeCountsByRegion();

        function computeStackedPositionsByRegion(countsByRegion) {
            // Build a square stacked block: all particles are assigned to a single NxN grid
            // so there are no "stray" particles outside the combined square.
            let chart = {
                cell: 3.0
            };

            let total = d3.sum(regionNames, r => (countsByRegion.get(r) || 0));
            if (!(total > 0)) total = 1;

            let margin = 20;
            let maxSide = Math.max(120, Math.min(svgWidth, svgHeight) - margin * 2);
            let neededN = Math.ceil(Math.sqrt(total));
            // If the particle count is high, shrink the cell size so everything fits in the SVG.
            // (This avoids leaving un-positioned particles at random locations.)
            let fittedCell = maxSide / neededN;
            if (!Number.isFinite(fittedCell) || fittedCell <= 0) fittedCell = 0.15;
            chart.cell = Math.min(chart.cell, fittedCell);
            let n = Math.max(1, neededN);
            let side = n * chart.cell;

            chart.width = side;
            chart.height = side;
            chart.x = Math.round((svgWidth - chart.width) / 2);
            chart.y = Math.round((svgHeight - chart.height) / 2);

            // Keep within bounds if layout changes.
            chart.x = Math.max(margin, Math.min(chart.x, svgWidth - chart.width - margin));
            chart.y = Math.max(margin, Math.min(chart.y, svgHeight - chart.height - margin));

            let out = new Map();
            let cols = n;

            // Order segments by count (largest at bottom feels more stable/legible).
            let ordered = regionNames
                .map(r => ({ r, c: countsByRegion.get(r) || 0 }))
                .filter(d => d.c > 0)
                .sort((a, b) => a.c - b.c);

            let cursor = 0;
            ordered.forEach(({ r, c }) => {
                let rng = mulberry32(fnv1a32(`${selectedYear}|${selectedMetric}|${r}|stacked`));
                let pts = new Array(c);
                for (let i = 0; i < c; i++) {
                    let idx = cursor + i;
                    let row = Math.floor(idx / cols);
                    let col = idx % cols;
                    let jitterX = (rng() - 0.5) * 0.8;
                    let jitterY = (rng() - 0.5) * 0.8;
                    pts[i] = [
                        chart.x + col * chart.cell + jitterX,
                        chart.y + row * chart.cell + jitterY
                    ];
                }

                out.set(r, pts);
                cursor += c;
            });

            return { pointsByRegion: out, chart, total };
        }

        function drawStackedPercentLegend(chart, countsByRegion) {
            if (!chart) return;
            let total = d3.sum(regionNames, r => (countsByRegion.get(r) || 0));
            if (!(total > 0)) return;

            let entries = regionNames
                .map(r => ({ region: r, count: countsByRegion.get(r) || 0 }))
                .filter(d => d.count > 0)
                .sort((a, b) => b.count - a.count)
                .map(d => ({ ...d, pct: (d.count / total) * 100 }));

            if (!entries.length) return;

            let pad = 10;
            let rowH = 18;
            let titleH = 18;
            let legendW = 320;
            let legendH = pad * 2 + titleH + entries.length * rowH;

            // Keep the share legend in the same consistent spot as the map legend (top-left).
            let x = 12;
            let y = Math.max(10, Math.min(12, svgHeight - legendH - 10));

            let g = chartLayer.append("g").attr("class", "stacked-legend");
            g.append("rect")
                .attr("x", x)
                .attr("y", y)
                .attr("width", legendW)
                .attr("height", legendH)
                .attr("rx", 10)
                .attr("fill", "rgba(255,255,255,0.92)");

            g.append("text")
                .attr("x", x + pad)
                .attr("y", y + pad + 12)
                .attr("fill", "#111827")
                .attr("font-family", "system-ui")
                .attr("font-size", 12)
                .attr("font-weight", 650)
                .text("Province share (%)");

            let fmt = d3.format(".1f");
            entries.forEach((d, i) => {
                let yy = y + pad + titleH + i * rowH + 12;
                g.append("rect")
                    .attr("x", x + pad)
                    .attr("y", yy - 10)
                    .attr("width", 10)
                    .attr("height", 10)
                    .attr("rx", 2)
                    .attr("fill", provinceColours[d.region] || "#6b7280");

                g.append("text")
                    .attr("x", x + pad + 16)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .text(d.region);

                g.append("text")
                    .attr("x", x + legendW - pad)
                    .attr("y", yy)
                    .attr("fill", "#111827")
                    .attr("font-family", "system-ui")
                    .attr("font-size", 12)
                    .attr("text-anchor", "end")
                    .text(`${fmt(d.pct)}%`);
            });
        }

        function computeStackedAreaPositionsByRegion(countsByRegion) {
            // Create a point cloud that fills the stacked-area shapes so particles can transition into it.
            let { pad, innerW, innerH, x, y, layers, years } = computeStackedAreaLayout(selectedMetric);
            let cell = 3.2;
            let pointsByRegion = new Map();
            regionNames.forEach(r => pointsByRegion.set(r, []));

            function lerp(a, b, t) { return a + (b - a) * t; }

            // Precompute segments for each layer between the three years.
            let yearNums = years.map(d => +d);
            let segXs = yearNums.map(yr => x(yr));

            for (let li = 0; li < layers.length; li++) {
                let layer = layers[li];
                let key = layer.key;
                // layer has 3 points: [y0,y1] at each year.
                for (let s = 0; s < layer.length - 1; s++) {
                    let xA = segXs[s];
                    let xB = segXs[s + 1];
                    let a0 = layer[s][0], a1 = layer[s][1];
                    let b0 = layer[s + 1][0], b1 = layer[s + 1][1];

                    let startX = Math.ceil(xA / cell) * cell;
                    let endX = xB;
                    for (let px = startX; px <= endX; px += cell) {
                        let t = (xB === xA) ? 0 : (px - xA) / (xB - xA);
                        t = Math.max(0, Math.min(1, t));
                        let y0v = lerp(a0, b0, t);
                        let y1v = lerp(a1, b1, t);

                        let top = y(y1v);
                        let bot = y(y0v);
                        // guard against degenerate segments
                        if (!(bot - top > 0.5)) continue;

                        let startY = Math.ceil(top / cell) * cell;
                        for (let py = startY; py <= bot; py += cell) {
                            // keep within chart frame
                            if (px < pad.left || px > pad.left + innerW) continue;
                            if (py < pad.top || py > pad.top + innerH) continue;
                            pointsByRegion.get(key)?.push([px, py]);
                        }
                    }
                }
            }

            // Shuffle each region deterministically and sample to the needed counts.
            let sampled = new Map();
            regionNames.forEach(region => {
                let pts = pointsByRegion.get(region) || [];
                if (!pts.length) {
                    sampled.set(region, []);
                    return;
                }
                let rng = mulberry32(fnv1a32(`${selectedMetric}|stacked_area|${region}`));
                for (let i = pts.length - 1; i > 0; i--) {
                    let j = (rng() * (i + 1)) | 0;
                    let tmp = pts[i]; pts[i] = pts[j]; pts[j] = tmp;
                }
                let need = Math.min(countsByRegion.get(region) || 0, pts.length);
                sampled.set(region, pts.slice(0, need));
            });

            return sampled;
        }

        function samplePointsWithReplacement(points, need, rng, jitter = 1.4) {
            if (need <= 0) return [];
            if (!points || !points.length) return [];
            if (points.length >= need) return points.slice(0, need);

            let out = new Array(need);
            for (let i = 0; i < need; i++) {
                let p = points[i % points.length];
                let dx = (rng() - 0.5) * 2 * jitter;
                let dy = (rng() - 0.5) * 2 * jitter;
                out[i] = [p[0] + dx, p[1] + dy];
            }
            return out;
        }

        function computePiePositionsByRegion(countsByRegion) {
            let legendW = 330;
            let cx = Math.max(legendW + 230, svgWidth * 0.62);
            let cy = svgHeight * 0.52;
            let outerR = Math.min(190, Math.min(svgWidth - cx - 18, svgHeight * 0.44));
            let innerR = outerR * 0.58;

            let cfgLocal = metricConfig[selectedMetric] || metricConfig.total_orders;
            let rows = regionNames.map(region => {
                let info = getProvinceData(selectedYear, region);
                let raw = cfgLocal.getValue(info);
                let v = Number.isFinite(raw) ? Math.max(0, raw) : 0;
                return { region, value: v };
            });

            let totalNeed = d3.sum(regionNames, r => (countsByRegion.get(r) || 0));
            if (!(totalNeed > 0)) totalNeed = 1;
            let donutArea = Math.PI * (outerR * outerR - innerR * innerR);
            let cell = Math.sqrt((donutArea / totalNeed) * 1.2);
            cell = Math.max(1.8, Math.min(4.8, cell));

            let pie = d3.pie()
                .sort(null)
                .value(d => d.value);

            let arcs = pie(rows);

            let candidateByRegion = new Map();
            regionNames.forEach(r => candidateByRegion.set(r, []));

            for (let i = 0; i < arcs.length; i++) {
                let a = arcs[i];
                let region = a.data.region;
                let need = countsByRegion.get(region) || 0;
                if (need <= 0 || !(a.endAngle > a.startAngle)) continue;

                let pts = candidateByRegion.get(region);
                for (let r = innerR + cell * 0.5; r <= outerR; r += cell) {
                    let arcLen = r * (a.endAngle - a.startAngle);
                    let nTheta = Math.max(1, Math.floor(arcLen / cell));
                    let step = (a.endAngle - a.startAngle) / nTheta;
                    for (let k = 0; k < nTheta; k++) {
                        let theta = a.startAngle + (k + 0.5) * step;
                        pts.push([
                            cx + r * Math.cos(theta),
                            cy + r * Math.sin(theta)
                        ]);
                    }
                }
            }

            let sampled = new Map();
            regionNames.forEach(region => {
                let need = countsByRegion.get(region) || 0;
                if (need <= 0) {
                    sampled.set(region, []);
                    return;
                }
                let pts = candidateByRegion.get(region) || [];
                let rng = mulberry32(fnv1a32(`${selectedMetric}|${selectedYear}|pie|${region}`));
                for (let i = pts.length - 1; i > 0; i--) {
                    let j = (rng() * (i + 1)) | 0;
                    let tmp = pts[i]; pts[i] = pts[j]; pts[j] = tmp;
                }
                sampled.set(region, samplePointsWithReplacement(pts, need, rng, 1.6));
            });

            return sampled;
        }

        function computeBarPositionsByRegion(countsByRegion) {
            let entries = getRegionMetricEntriesForYear(selectedYear, selectedMetric)
                .filter(d => d.value > 0)
                .sort((a, b) => b.value - a.value);

            let pad = { left: 210, right: 22, top: 34, bottom: 34 };
            let innerW = svgWidth - pad.left - pad.right;
            let innerH = svgHeight - pad.top - pad.bottom;

            let maxV = d3.max(entries, d => d.value) || 1;
            let x = d3.scaleLinear().domain([0, maxV]).nice().range([pad.left, pad.left + innerW]);
            let y = d3.scaleBand().domain(entries.map(d => d.region)).range([pad.top, pad.top + innerH]).padding(0.22);

            let sampled = new Map();
            regionNames.forEach(region => sampled.set(region, []));

            entries.forEach(d => {
                let region = d.region;
                let need = countsByRegion.get(region) || 0;
                if (need <= 0) return;

                let barW = Math.max(0, x(d.value) - pad.left);
                let y0 = y(region);
                if (y0 == null) return;
                let barH = y.bandwidth();
                if (!(barW > 0 && barH > 0)) return;

                let area = barW * barH;
                let cell = Math.sqrt((area / need) * 1.35);
                cell = Math.max(1.7, Math.min(4.6, cell));

                let pts = [];
                for (let px = pad.left + cell * 0.5; px <= pad.left + barW - cell * 0.35; px += cell) {
                    for (let py = y0 + cell * 0.5; py <= y0 + barH - cell * 0.35; py += cell) {
                        pts.push([px, py]);
                    }
                }

                let rng = mulberry32(fnv1a32(`${selectedMetric}|${selectedYear}|bar|${region}`));
                for (let i = pts.length - 1; i > 0; i--) {
                    let j = (rng() * (i + 1)) | 0;
                    let tmp = pts[i]; pts[i] = pts[j]; pts[j] = tmp;
                }
                sampled.set(region, samplePointsWithReplacement(pts, need, rng, 1.1));
            });

            return sampled;
        }

        regionNames.forEach(region => {
            let info = getProvinceData(selectedYear, region);
            let value = cfg.getValue(info);
            if (!Number.isFinite(value)) value = 0;

            let count = countsByRegion.get(region) ?? 0;
            if (count <= 0) return;

            if (viewMode === viewModes.map) {
                let maskPoints = pointsForRegion(region, count);
                if (maskPoints) {
                    for (let i = 0; i < count; i++) {
                        let pt = maskPoints[i];
                        particles.push({
                            id: `${region}-${i}`,
                            x: pt[0],
                            y: pt[1],
                            colour: provinceColours[region] || "#555",
                            location: region
                        });
                    }
                    return;
                }
            }

            for (let i = 0; i < count; i++) {
                let pt = (viewMode === viewModes.map)
                    ? randomPointInShape(provinceShapes[region])
                    : [Math.random() * svgWidth, Math.random() * svgHeight];
                if (!pt) break;
                particles.push({
                    id: `${region}-${i}`,
                    x: pt[0],
                    y: pt[1],
                    colour: provinceColours[region] || "#555",
                    location: region
                });
            }
        });

        if (viewMode === viewModes.stacked) {
            let stacked = computeStackedPositionsByRegion(countsByRegion);
            let stackedPts = stacked.pointsByRegion;
            drawStackedPercentLegend(stacked.chart, countsByRegion);
            let offsetByRegion = new Map();
            regionNames.forEach(r => offsetByRegion.set(r, 0));
            for (let i = 0; i < particles.length; i++) {
                let p = particles[i];
                let pts = stackedPts.get(p.location);
                if (!pts || !pts.length) continue;
                let idx = offsetByRegion.get(p.location) || 0;
                if (idx >= pts.length) continue;
                let pt = pts[idx];
                offsetByRegion.set(p.location, idx + 1);
                p.x = pt[0];
                p.y = pt[1];
            }
        }

        if (viewMode === viewModes.stacked_area) {
            let areaPts = computeStackedAreaPositionsByRegion(countsByRegion);
            let offsetByRegion = new Map();
            regionNames.forEach(r => offsetByRegion.set(r, 0));
            for (let i = 0; i < particles.length; i++) {
                let p = particles[i];
                let pts = areaPts.get(p.location);
                if (!pts || !pts.length) continue;
                let idx = offsetByRegion.get(p.location) || 0;
                if (idx >= pts.length) continue;
                let pt = pts[idx];
                offsetByRegion.set(p.location, idx + 1);
                p.x = pt[0];
                p.y = pt[1];
            }
        }

        if (viewMode === viewModes.pie) {
            let piePts = computePiePositionsByRegion(countsByRegion);
            let offsetByRegion = new Map();
            regionNames.forEach(r => offsetByRegion.set(r, 0));
            for (let i = 0; i < particles.length; i++) {
                let p = particles[i];
                let pts = piePts.get(p.location);
                if (!pts || !pts.length) continue;
                let idx = offsetByRegion.get(p.location) || 0;
                if (idx >= pts.length) continue;
                let pt = pts[idx];
                offsetByRegion.set(p.location, idx + 1);
                p.x = pt[0];
                p.y = pt[1];
            }
        }

        if (viewMode === viewModes.bar) {
            let barPts = computeBarPositionsByRegion(countsByRegion);
            let offsetByRegion = new Map();
            regionNames.forEach(r => offsetByRegion.set(r, 0));
            for (let i = 0; i < particles.length; i++) {
                let p = particles[i];
                let pts = barPts.get(p.location);
                if (!pts || !pts.length) continue;
                let idx = offsetByRegion.get(p.location) || 0;
                if (idx >= pts.length) continue;
                let pt = pts[idx];
                offsetByRegion.set(p.location, idx + 1);
                p.x = pt[0];
                p.y = pt[1];
            }
        }

        relocateIsolatedParticles(particles);

        let transition = d3.transition().duration(650).ease(d3.easeCubicOut);

        function highlightProvince(nameOrNull) {
            if (mapMask.ready) return;
            provinceLayer.selectAll(".province")
                .attr("fill-opacity", p => (nameOrNull && p === nameOrNull) ? 0.18 : 0.08)
                .attr("stroke-width", p => (nameOrNull && p === nameOrNull) ? 2 : 1);
        }

        function showTooltip(event, regionName) {
            let info = getProvinceData(selectedYear, regionName);
            highlightProvince(regionName);
            tooltip
                .style("opacity", 1)
                .html(`
                        <strong>${regionName} (${selectedYear})</strong><br/>
                        Orders: ${formatMetricValue("total_orders", info.orders)}<br/>
                        Avg Orders: ${formatMetricValue("avg_orders", info.avgOrders)}<br/>
                        Total Value: ${formatMetricValue("value_orders", info.value)}<br/>
                        Avg Value/Person: ${formatMetricValue("avg_value_per_person", info.avgValue)}
                    `)
                .style("left", (event.pageX + 10) + "px")
                .style("top", (event.pageY + 10) + "px");
        }

        function hideTooltip() {
            highlightProvince(null);
            tooltip.style("opacity", 0);
        }

        function randomStart() {
            return [Math.random() * svgWidth, Math.random() * svgHeight];
        }

        let circles = particleLayer.selectAll("circle.particle")
            .data(particles, d => d.id)
            .join(
                enter => {
                    return enter
                        .append("circle")
                        .attr("class", "particle")
                        .each(function() {
                            let [sx, sy] = randomStart();
                            d3.select(this).attr("cx", sx).attr("cy", sy);
                        })
                        .attr("r", 0)
                        .attr("fill", d => d.colour)
                        .attr("opacity", 0)
                        .call(sel => sel.transition(transition)
                            .attr("cx", d => d.x)
                            .attr("cy", d => d.y)
                            .attr("r", 1.3)
                            .attr("opacity", 0.65)
                        );
                },
                update => {
                    return update.call(sel => sel.transition(transition)
                        .attr("cx", d => d.x)
                        .attr("cy", d => d.y)
                        .attr("fill", d => d.colour)
                        .attr("opacity", 0.65)
                    );
                },
                exit => {
                    return exit.call(sel => sel.transition()
                        .duration(400)
                        .attr("r", 0)
                        .attr("opacity", 0)
                        .remove()
                    );
                }
            );

        circles
            .on("mouseover", function(event, d) {
                showTooltip(event, d.location);
            })
            .on("mousemove", function(event) {
                tooltip
                    .style("left", (event.pageX + 10) + "px")
                    .style("top", (event.pageY + 10) + "px");
            })
            .on("mouseout", function() {
                // Let the proximity handler decide when to hide to avoid requiring pixel-perfect hover.
            });

        // Proximity tooltip: show province tooltip when near the nearest particle (no need to hover exactly).
        svg.on("mousemove.proximity", null).on("mouseleave.proximity", null);
        if (particles.length) {
            let delaunay = d3.Delaunay.from(particles, d => d.x, d => d.y);
            let lastRegion = null;
            // Hysteresis reduces flicker: show when close, hide when clearly away.
            let showDist = 28; // px
            let hideDist = 44; // px
            let showDist2 = showDist * showDist;
            let hideDist2 = hideDist * hideDist;

            svg.on("mousemove.proximity", function(event) {
                let [mx, my] = d3.pointer(event, svg.node());
                // `particles` are in data coords; find nearest using inverted zoom transform.
                let [dxp, dyp] = zoomTransform.invert([mx, my]);
                let i = delaunay.find(dxp, dyp);
                if (i == null) {
                    if (lastRegion != null) hideTooltip();
                    lastRegion = null;
                    return;
                }

                let p = particles[i];
                // Compare distance in *screen* pixels so thresholds feel consistent while zoomed.
                let sx = zoomTransform.applyX(p.x);
                let sy = zoomTransform.applyY(p.y);
                let dx = sx - mx;
                let dy = sy - my;
                let d2 = dx * dx + dy * dy;
                if (lastRegion == null) {
                    if (d2 > showDist2) return;
                } else if (d2 > hideDist2) {
                    hideTooltip();
                    lastRegion = null;
                    return;
                }

                if (p.location !== lastRegion) {
                    showTooltip(event, p.location);
                    lastRegion = p.location;
                } else {
                    tooltip
                        .style("left", (event.pageX + 10) + "px")
                        .style("top", (event.pageY + 10) + "px");
                }
            });

            svg.on("mouseleave.proximity", function() {
                if (lastRegion != null) hideTooltip();
                lastRegion = null;
            });
        }
    }

    function render() {
        let year = d3.select("#year-select").property("value") || "2005";
        let metric = d3.select("#metric-select").property("value") || "total_orders";
        let viewMode = d3.select("#view-select").property("value") || viewModes.map;

        if (viewMode === viewModes.map) {
            if (zoomToggleBtn) zoomToggleBtn.style.display = "";
            if (zoomEnabled) enableZoom();
            else disableZoom();
        } else {
            if (zoomToggleBtn) zoomToggleBtn.style.display = "none";
            disableZoomAndReset();
        }

        chartLayer.selectAll("*").remove();
        bgLayer.selectAll("*").remove();
        provinceLayer.selectAll("*").remove();
        labelLayer.selectAll("*").remove();
        uiLayer.selectAll("*").remove();
        hideChartTooltip();
        updateFactsPanel(year, metric);
        if (viewMode === viewModes.stacked_area) drawStackedAreaChart(year, metric);
        if (viewMode === viewModes.pie) drawPieChart(year, metric);
        if (viewMode === viewModes.bar) drawBarChart(year, metric);
        if (viewMode === viewModes.map) drawMapLegend();

        drawProvinces(viewMode);
        drawLabels(viewMode);
        drawParticles(year, metric, viewMode);
    }

    d3.select("#view-select").on("change", render);

    function syncZoomToggleUI() {
        if (!zoomToggleBtn) return;
        zoomToggleBtn.setAttribute("aria-pressed", zoomEnabled ? "true" : "false");
        zoomToggleBtn.textContent = zoomEnabled ? "Zoom: On" : "Zoom: Off";
    }

    if (zoomToggleBtn) {
        zoomToggleBtn.addEventListener("click", () => {
            zoomEnabled = !zoomEnabled;
            syncZoomToggleUI();
            render();
        });
        syncZoomToggleUI();
    }

    // Timeline (scrubber) under the map
    let timelineYears = ["2005", "2007", "2009"];
    let yearSelect = document.getElementById("year-select");
    let yearTimeline = document.getElementById("year-timeline");
    let yearPrev = document.getElementById("year-prev");
    let yearNext = document.getElementById("year-next");

    function setYearValue(year) {
        if (yearSelect) yearSelect.value = year;
        if (yearTimeline) yearTimeline.value = String(Math.max(0, timelineYears.indexOf(year)));
        render();
    }

    function syncTimelineFromSelect() {
        if (!yearSelect || !yearTimeline) return;
        let idx = timelineYears.indexOf(yearSelect.value);
        yearTimeline.value = String(Math.max(0, idx));
    }

    if (yearTimeline) {
        yearTimeline.addEventListener("input", () => {
            let idx = Math.max(0, Math.min(timelineYears.length - 1, parseInt(yearTimeline.value, 10) || 0));
            setYearValue(timelineYears[idx]);
        });

        // Allow scrolling the mouse wheel to change years when hovering the scrubber.
        yearTimeline.addEventListener("wheel", (e) => {
            e.preventDefault();
            let dir = e.deltaY > 0 ? 1 : -1;
            let cur = yearSelect ? timelineYears.indexOf(yearSelect.value) : (parseInt(yearTimeline.value, 10) || 0);
            let next = Math.max(0, Math.min(timelineYears.length - 1, cur + dir));
            setYearValue(timelineYears[next]);
        }, { passive: false });
    }

    document.querySelectorAll(".timeline-label").forEach(el => {
        el.addEventListener("click", () => {
            let year = el.getAttribute("data-year");
            if (year && timelineYears.includes(year)) setYearValue(year);
        });
    });

    if (yearPrev) yearPrev.addEventListener("click", () => {
        let cur = yearSelect ? timelineYears.indexOf(yearSelect.value) : 0;
        setYearValue(timelineYears[Math.max(0, cur - 1)]);
    });
    if (yearNext) yearNext.addEventListener("click", () => {
        let cur = yearSelect ? timelineYears.indexOf(yearSelect.value) : 0;
        setYearValue(timelineYears[Math.min(timelineYears.length - 1, cur + 1)]);
    });

    d3.select("#year-select").on("change", () => {
        syncTimelineFromSelect();
        render();
    });
    d3.select("#metric-select").on("change", render);

    // Ensure timeline matches the default dropdown value on first load.
    syncTimelineFromSelect();

    initMapMask().finally(() => render());
})
