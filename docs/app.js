import * as maplibregl from "https://unpkg.com/maplibre-gl@6.15.0/dist/maplibre-gl.mjs";

const { createApp, nextTick, markRaw } = Vue;

createApp({
  data() {
    return {
      items: [], summary: null, page: "home", selectedIndex: null, selectedItem: null,
      query: "", searchField: "all", categoryFilters: [], versionFilters: [],
      resourceTypeFilters: [], publisherFilters: [], providerFilters: [], placeFilters: [], sortBy: "title", resultPage: 1, resultPageSize: 20,
      assetQuery: "", assetSubtype: "", assetSource: "", assetRole: "", assetPage: 1, assetPageSize: 10,
      loading: true, loadError: "", mapInstance: null, mapReady: false, mapLoadPromise: null, toastMessage: "", toastTimeout: null
    };
  },
  computed: {
    categories() {
      return [...new Set(this.items.map((item) => item.category || "Uncategorized"))].sort((a, b) => a.localeCompare(b));
    },
    resourceTypes() {
      return [...new Set(this.items.map((item) => item.resource_type).filter(Boolean))].sort();
    },
    
    publishers() {
      return [...new Set(
        this.items.flatMap(item => item.publishers || [])
      )].filter(Boolean).sort();
    },

    providers() {
      return [...new Set(
        this.items.flatMap(item => item.providers || [])
      )].filter(Boolean).sort();
    },

    places() {
      return [...new Set(
        this.items.flatMap(item => item.place || [])
      )].filter(Boolean).sort();
    },
    totalAssets() { return this.items.reduce((sum, item) => sum + Number(item.asset_count || 0), 0); },
    currentCount() { return this.items.filter((item) => !item.version_status || item.version_status === "unknown" || item.version_status === "current").length; },
    
    categoryCards() {
      return this.categories.map((name) => {
        const rows = this.items.filter(
          (item) => (item.category || "Uncategorized") === name
        );

        return {
          name,
          count: rows.length,

          assets: rows.reduce(
            (sum, item) => sum + Number(item.asset_count || 0),
            0
          ),

          current: rows.filter(
            (item) => !item.version_status || item.version_status === "unknown" || item.version_status === "current"
          ).length,

          preview:
            rows.find((item) => item.description)?.description?.slice(0, 105) || ""
        };
      }).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

    
    },

    
    filteredItems() {
      const needle = this.query.trim().toLocaleLowerCase();

      const fields = {
        title: ["title"],
        doi: ["dois"],
        tags: ["keywords", "tags"],
        description: ["description"],
        authors: ["authors"],
        publishers: ["publishers"],
        providers: ["providers"],
        place: ["place"],
        purpose: ["purpose"],
        citation: ["citation"],
        version: ["version"],
      };

      let rows = this.items.filter((item) => {
        // Search
        const selectedFields = this.searchField === "all"
          ? Object.keys(fields)
          : [this.searchField];

        const haystack = selectedFields
          .flatMap(field => fields[field] || [])
          .map(key => this.valueText(item[key]))
          .join(" ")
          .toLocaleLowerCase();

        const matchesSearch = !needle || haystack.includes(needle);

        // Publisher
        const matchesPublisher =
          !this.publisherFilters.length ||
          this.publisherFilters.some(value =>
            (item.publishers || []).includes(value)
          );

        // Provider
        const matchesProvider =
          !this.providerFilters.length ||
          this.providerFilters.some(value =>
            (item.providers || []).includes(value)
          );

        // Geographic coverage
        const matchesPlace =
          !this.placeFilters.length ||
          this.placeFilters.some(value =>
            (item.place || []).includes(value)
          );

        // Existing filters
        const matchesCategory =
          !this.categoryFilters.length ||
          this.categoryFilters.includes(item.category || "Uncategorized");

        const matchesVersion =
          !this.versionFilters.length ||
          this.versionFilters.includes(
            String(item.version_status || "not specified").trim().toLowerCase()
          );

        const matchesResourceType =
          !this.resourceTypeFilters.length ||
          this.resourceTypeFilters.includes(item.resource_type);

        return matchesSearch &&
          matchesPublisher &&
          matchesProvider &&
          matchesPlace &&
          matchesCategory &&
          matchesVersion &&
          matchesResourceType;
      });

      rows = [...rows];

      if (this.sortBy === "assets") {
        rows.sort((a, b) =>
          Number(b.asset_count || 0) - Number(a.asset_count || 0) ||
          a.title.localeCompare(b.title)
        );
      } else if (this.sortBy === "category") {
        rows.sort((a, b) =>
          a.category.localeCompare(b.category) ||
          a.title.localeCompare(b.title)
        );
      } else if (this.sortBy === "date") {
        rows.sort((a, b) =>
          this.dateValue(b.publication_date) - this.dateValue(a.publication_date) ||
          a.title.localeCompare(b.title)
        );
      } else {
        rows.sort((a, b) => a.title.localeCompare(b.title));
      }

      return rows;
    },

    resultPageCount() { return Math.max(1, Math.ceil(this.filteredItems.length / this.resultPageSize)); },
    pagedItems() { const start = (this.resultPage - 1) * this.resultPageSize; return this.filteredItems.slice(start, start + this.resultPageSize); },
    categoryItems() {
      if (!this.categoryFilters.length) return [];

      return this.items.filter((item) =>
        this.categoryFilters.includes(item.category || "Uncategorized")
      );
    },
    categoryAssets() { return this.categoryItems.reduce((sum, item) => sum + Number(item.asset_count || 0), 0); },
    categoryCurrent() {
      return this.categoryItems.filter((item) => {
        const status = String(item.version_status ?? "")
          .trim()
          .toLowerCase();

        return (
          !status ||
          status === "current" ||
          status === "unknown" ||
          status === "not specified"
        );
      }).length;
    },
    categoryTopics() {
      const counts = new Map();
      this.categoryItems.forEach((item) => (item.keywords || item.tags || []).forEach((tag) => counts.set(tag, (counts.get(tag) || 0) + 1)));
      return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).slice(0, 8).map(([tag]) => tag);
    },
    selectedProps() { return this.selectedItem?.properties || {}; },
    versionStatus() { return this.selectedProps["ciesin:version_status"] || this.selectedIndex?.version_status || "unknown"; },
    versionLabel() { return this.versionStatus === "current" ? "Current version" : this.versionStatus === "not_current" ? "Not current" : "Version status not specified"; },
    selectedAssets() {
      return Object.entries(this.selectedItem?.assets || {}).map(([key, asset]) => ({ key, ...asset,
        source: asset["ciesin:source"] || "", subtype: asset["ciesin:asset_subtype"] || "Unclassified",
        roles: Array.isArray(asset.roles) ? asset.roles : [] }));
    },
    assetSubtypes() { return [...new Set(this.selectedAssets.map((asset) => asset.subtype))].sort(); },
    assetSources() { return [...new Set(this.selectedAssets.map((asset) => asset.source).filter(Boolean))].sort(); },
    assetRoles() { return [...new Set(this.selectedAssets.flatMap((asset) => asset.roles).filter(Boolean))].sort(); },
    filteredAssets() {
      const q = this.assetQuery.trim().toLocaleLowerCase();
      return this.selectedAssets.filter((asset) => {
        const text = [asset.title, asset.href, asset.description, asset.type, asset.source, asset.subtype,
          asset["ciesin:source_identifier"], ...(asset.roles || [])].join(" ").toLocaleLowerCase();
        return (!q || text.includes(q)) && (!this.assetSource || asset.source === this.assetSource)
          && (!this.assetSubtype || asset.subtype === this.assetSubtype)
          && (!this.assetRole || asset.roles.includes(this.assetRole));
      });
    },
    assetPageCount() { return Math.max(1, Math.ceil(this.filteredAssets.length / this.assetPageSize)); },
    pagedAssets() { const start = (this.assetPage - 1) * this.assetPageSize; return this.filteredAssets.slice(start, start + this.assetPageSize); },
    hasBbox() { return Array.isArray(this.selectedItem?.bbox) && this.selectedItem.bbox.length === 4 && this.selectedItem.bbox.every(Number.isFinite); },
    metadataRows() {
      const p = this.selectedProps;
      const rows = [
        ["Category", p["ciesin:category"] || this.selectedIndex?.category],
        ["Resource type", p["ciesin:resource_type"] || this.selectedIndex?.resource_type],
        ["Version status", this.versionLabel], ["Version", p.version || this.selectedIndex?.version],
        ["DOIs", p["ciesin:dois"] || this.selectedIndex?.dois],
        ["Publication date", p["ciesin:publication_date"] || this.selectedIndex?.publication_date],
        ["Data coverage", this.coverageLabel], ["Bounding box (W, S, E, N)", this.hasBbox ? this.selectedItem.bbox : ""],
        ["Tags", p["ciesin:tags"] || p.keywords], ["License", p.license],
        ["Providers", p.providers], ["Authors", p["ciesin:authors"]], ["Publishers", p["ciesin:publishers"]],
        ["Place", p["ciesin:place"]], ["Citation", p["sci:citation"] || this.selectedIndex?.citation],
        ["Purpose", p["ciesin:purpose"] || this.selectedIndex?.purpose],
        ["Additional resources", p["ciesin:additional_resources"] || this.selectedIndex?.additional_resources],
        ["Metadata sources", p["ciesin:metadata_sources"]], ["XML metadata location", p["ciesin:inventory_xml_location"]],
      ];
      return rows.map(([label, value]) => ({ label, value: this.valueText(value) })).filter((row) => row.value);
    },
    coverageLabel() {
      const p = this.selectedProps;
      const start = p.start_datetime || p["ciesin:temporal_start"];
      const end = p.end_datetime || p["ciesin:temporal_end"];
      if (start && end) return `${this.formatDate(start)} — ${this.formatDate(end)}`;
      return start ? `From ${this.formatDate(start)}` : end ? `Through ${this.formatDate(end)}` : this.formatDate(p.datetime);
    },
    selectedCategoryLabel() {
      return this.categoryFilters.join(", ");
    },
    publishers() {
      return [...new Set(
        this.items.flatMap(item => item.publishers || [])
      )].filter(Boolean).sort();
    },

    providers() {
      return [...new Set(
        this.items.flatMap(item => item.providers || [])
      )].filter(Boolean).sort();
    },

    places() {
      return [...new Set(
        this.items.flatMap(item => item.place || [])
      )].filter(Boolean).sort();
    },
  },
  watch: {
    query() { this.resultPage = 1; },
    categoryFilters() { this.resultPage = 1; },
    versionFilters() { this.resultPage = 1; },
    resourceTypeFilters() { this.resultPage = 1; },
    publisherFilters() { this.resultPage = 1; },
    providerFilters() { this.resultPage = 1; },
    placeFilters() { this.resultPage = 1; },
    assetQuery() { this.assetPage = 1; }, assetSubtype() { this.assetPage = 1; }, assetSource() { this.assetPage = 1; }, assetRole() { this.assetPage = 1; },
    pagedAssets() { if (this.assetPage > this.assetPageCount) this.assetPage = this.assetPageCount; },
    async selectedItem() {
          if (this.page === "item" && this.selectedItem) {
            this.updateMap();
          }
      },
  },
  async mounted() {
    window.addEventListener("popstate", this.handleLocationChange);

    try {
      const response = await fetch("./stac/items-index.json");
      if (!response.ok) {
        throw new Error(`Could not read catalog index (HTTP ${response.status}).`);
      }

      this.items = await response.json();

      const summaryResponse = await fetch("./stac/summary.json");
      if (summaryResponse.ok) {
        this.summary = await summaryResponse.json();
      }

      // Open a dataset directly when its ID is in the URL.
      const params = new URLSearchParams(window.location.search);
      const itemId = params.get("item");

      if (itemId) {
        const match = this.items.find(
          (item) => String(item.id) === itemId
        );

        if (match) {
          await this.openItem(match, { updateUrl: false });
        } else {
          this.loadError = `Dataset "${itemId}" was not found in the catalog.`;
        }
      }
    } catch (error) {
      this.loadError = `${error.message} Build the STAC data first, then serve this docs folder over HTTP.`;
    } finally {
      this.loading = false;
    }
  },
  beforeUnmount() {
    window.removeEventListener("popstate", this.handleLocationChange);

    if (this.mapInstance) {
      this.mapInstance.remove();
      this.mapInstance = null;
    }
  },
  methods: {
    valueText(value) {
      if (value === null || value === undefined || value === "") return "";
      if (Array.isArray(value)) return value.map((entry) => this.valueText(entry)).filter(Boolean).join("; ");
      if (typeof value === "object") return Object.values(value).map((entry) => this.valueText(entry)).filter(Boolean).join("; ");
      return String(value);
    },
    dateValue(value) { const time = Date.parse(value || ""); return Number.isFinite(time) ? time : 0; },
    formatDate(value) { if (!value) return ""; const date = new Date(value); return Number.isNaN(date.getTime()) ? String(value) : new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeZone: "UTC" }).format(date); },
    formatBytes(value) { const bytes = Number(value); if (!Number.isFinite(bytes)) return ""; if (bytes < 1024) return `${bytes} B`; const units = ["KB", "MB", "GB", "TB"]; let size = bytes / 1024; let unit = 0; while (size >= 1024 && unit < units.length - 1) { size /= 1024; unit++; } return `${size.toFixed(1)} ${units[unit]}`; },
    doiHref(doi) { return `https://doi.org/${String(doi || "").replace(/^https?:\/\/(dx\.)?doi\.org\//i, "")}`; },
    safeHref(value) { try { const url = new URL(value); return ["https:", "http:"].includes(url.protocol) ? url.href : "#"; } catch { return "#"; } },
    categoryDescription(category) {
      const count = this.items.filter(
        (item) => (item.category || "Uncategorized") === category
      ).length;

      return `This collection contains ${count.toLocaleString()} datasets assigned to “${category}” in the CIESIN inventory. Browse the records below to compare descriptions, coverage, version status, and available assets.`;
    },
    openCategory(category) { this.categoryFilter = category; this.page = "home"; this.resultPage = 1; window.scrollTo({ top: 0, behavior: "smooth" }); },
    clearFilters() {
      this.query = "";
      this.searchField = "all";
      this.categoryFilters = [];
      this.versionFilters = [];
      this.resourceTypeFilters = [];
      this.sortBy = "title";
      this.resultPage = 1;
      this.publisherFilters = [];
      this.providerFilters = [];
      this.placeFilters = [];
    },
    goHome() {
      const url = new URL(window.location.href);
      url.searchParams.delete("item");
      window.history.replaceState({}, "", url);

      this.page = "home";
      this.selectedItem = null;
      this.selectedIndex = null;
      this.categoryFilters = [];
      this.versionFilters = [];
      this.resourceTypeFilters = [];
      this.mapReady = false;

      window.scrollTo({ top: 0, behavior: "smooth" });
    },

    goBack() {
      this.goHome();
    },
    async openItem(index, { updateUrl = true } = {}) {
      this.selectedIndex = index;
      this.selectedItem = null;
      this.page = "item";
      this.assetPage = 1;

      this.assetQuery = "";
      this.assetSubtype = "";
      this.assetSource = "";
      this.assetRole = "";
      this.mapReady = false;

      try {
        const response = await fetch(`./${index.item_href}`);

        if (!response.ok) {
          throw new Error(`Could not read item record (HTTP ${response.status}).`);
        }

        this.selectedItem = await response.json();

        // Add the dataset ID to the URL without reloading the page.
        if (updateUrl && index.id) {
          const url = new URL(window.location.href);
          url.searchParams.set("item", index.id);

          if (url.href !== window.location.href) {
            window.history.pushState({}, "", url);
          }
        }

        window.scrollTo({ top: 0, behavior: "smooth" });
      } catch (error) {
        this.loadError = error.message;
      }
    },
    

        
    async updateMap() {
      await nextTick();

      const container = document.getElementById("coverage-map");

      if (!container || !this.selectedItem) {
        console.warn("Map container or selected item is missing.");
        return;
      }

      // Create the map if it doesn't exist or belongs to an old container.
      if (
        !this.mapInstance ||
        this.mapInstance.getContainer() !== container
      ) {
        if (this.mapInstance) {
          this.mapInstance.remove();
          this.mapInstance = null;
        }

        console.log("Creating map for:", this.selectedItem.id);

        const map = markRaw(
          new maplibregl.Map({
            container,
            style: "https://tiles.openfreemap.org/styles/liberty",
            center: [0, 0],
            zoom: 1,
            attributionControl: true,
          })
        );

        this.mapInstance = map;

        map.addControl(
          new maplibregl.NavigationControl({ showCompass: false }),
          "top-right"
        );

        map.on("error", (event) => {
          console.error("MapLibre error:", event.error || event);
        });

        map.once("load", () => {
          console.log("Map style loaded.");

          if (this.mapInstance !== map) return;

          this.addCoverageLayers();
          this.paintCoverage();
          map.resize();
          this.mapReady = true;
        });

        return;
      }

      const map = this.mapInstance;

      if (map.isStyleLoaded()) {
        this.addCoverageLayers();
        this.paintCoverage();
        map.resize();
        this.mapReady = true;
      } else {
        map.once("load", () => {
          if (this.mapInstance !== map) return;

          this.addCoverageLayers();
          this.paintCoverage();
          map.resize();
          this.mapReady = true;
        });
      }
    },



    
    addCoverageLayers() {
      const map = this.mapInstance;
      if (!map || !map.isStyleLoaded()) return;

      if (!map.getSource("coverage-box")) {
        map.addSource("coverage-box", {
          type: "geojson",
          data: {
            type: "FeatureCollection",
            features: []
          }
        });
      }

      if (!map.getLayer("coverage-fill")) {
        map.addLayer({
          id: "coverage-fill",
          type: "fill",
          source: "coverage-box",
          paint: {
            "fill-color": "#187a70",
            "fill-opacity": 0.2
          }
        });
      }

      if (!map.getLayer("coverage-outline")) {
        map.addLayer({
          id: "coverage-outline",
          type: "line",
          source: "coverage-box",
          paint: {
            "line-color": "#075b55",
            "line-width": 2.5
          }
        });
      }
    },

    
    paintCoverage() {
      const map = this.mapInstance;
      if (!map || !map.getSource("coverage-box")) return;

      const item = this.selectedItem;
      if (!item) return;

      // Support either a bbox array or separate coordinate fields.
      const bbox = Array.isArray(item.bbox)
        ? item.bbox
        : [item.west, item.south, item.east, item.north];

      if (
        bbox.length !== 4 ||
        !bbox.every((value) => value !== null && value !== "" && Number.isFinite(Number(value)))
      ) {
        map.getSource("coverage-box").setData({
          type: "FeatureCollection",
          features: []
        });
        return;
      }

      const [west, south, east, north] = bbox.map(Number);

      // Draw the bounding box as a rectangle.
      const geometry = {
        type: "Polygon",
        coordinates: [[
          [west, south],
          [east, south],
          [east, north],
          [west, north],
          [west, south]
        ]]
      };

      map.getSource("coverage-box").setData({
        type: "FeatureCollection",
        features: [{
          type: "Feature",
          properties: {},
          geometry
        }]
      });

      map.fitBounds(
        [[west, south], [east, north]],
        {
          padding: 48,
          maxZoom: 5,
          duration: 700
        }
      );

      map.resize();
    },

    handleLocationChange() {
      const params = new URLSearchParams(window.location.search);
      const itemId = params.get("item");

      if (itemId) {
        const match = this.items.find(
          (item) => String(item.id) === itemId
        );

        if (match) {
          this.openItem(match, { updateUrl: false });
        }
      } else {
        // URL no longer identifies an item: return to the catalog.
        this.page = "home";
        this.selectedItem = null;
        this.selectedIndex = null;
        this.mapReady = false;
      }
    },

    itemShareUrl() {
      if (!this.selectedIndex?.id) return "";

      const url = new URL(window.location.href);
      url.searchParams.set("item", this.selectedIndex.id);

      return url.href;
    },

    async copyItemLink() {
      const url = this.itemShareUrl();

      if (!url) return;

      try {
        await navigator.clipboard.writeText(url);
        this.showToast("Dataset link copied to clipboard.");
      } catch (error) {
        console.error("Could not copy dataset link:", error);
      }
    },

    showToast(message) {
      this.toastMessage = message;

      clearTimeout(this.toastTimeout);

      this.toastTimeout = setTimeout(() => {
        this.toastMessage = "";
      }, 2000);
    },
  },
}).mount("#app");
