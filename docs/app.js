const { createApp, nextTick } = Vue;

createApp({
  data() {
    return {
      items: [],
      summary: null,
      selected: null,
      selectedItem: null,
      query: "",
      categoryFilter: "",
      statusFilter: "",
      sourceFilter: "",
      roleFilter: "",
      sortKey: "category",
      showAllCategories: false,
      page: 1,
      pageSize: 24,
      assetQuery: "",
      assetSourceFilter: "",
      assetRoleFilter: "",
      loadingMessage: "Reading stac/items-index.json",
    };
  },
  computed: {
    categories() {
      return [...new Set(this.items.map((item) => item.category || "Uncategorized"))].sort();
    },
    statuses() {
      return [...new Set(this.items.map((item) => item.item_status || item.match_status).filter(Boolean))].sort();
    },
    assetSources() {
      const sources = new Set();
      this.items.forEach((item) => Object.keys(item.source_counts || {}).forEach((source) => sources.add(source)));
      return [...sources].sort();
    },
    assetRoles() {
      const roles = new Set();
      this.items.forEach((item) => Object.keys(item.role_counts || {}).forEach((role) => roles.add(role)));
      return [...roles].sort();
    },
    allAssets() {
      return this.items.reduce((sum, item) => sum + Number(item.asset_count || 0), 0);
    },
    collectionCount() {
      return this.categories.length;
    },
    unmatchedAssets() {
      return Number(this.summary?.asset_rows_without_item_id_or_href || 0);
    },
    filteredItems() {
      const q = this.query.toLowerCase();
      let rows = this.items.filter((item) => {
        const haystack = [
          item.title,
          item.description,
          item.category,
          item.nasa_doi,
          item.harvard_doi,
          item.sedac_doi,
          item.collection,
          ...(item.keywords || []),
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        const itemStatus = item.item_status || item.match_status || "";
        const hasSource = !this.sourceFilter || (item.source_counts || {})[this.sourceFilter];
        const hasRole = !this.roleFilter || (item.role_counts || {})[this.roleFilter];
        return (
          (!q || haystack.includes(q)) &&
          (!this.categoryFilter || item.category === this.categoryFilter) &&
          (!this.statusFilter || itemStatus === this.statusFilter) &&
          hasSource &&
          hasRole
        );
      });

      rows = [...rows];
      if (this.sortKey === "assets") {
        rows.sort((a, b) => Number(b.asset_count || 0) - Number(a.asset_count || 0) || a.title.localeCompare(b.title));
      } else if (this.sortKey === "updated") {
        rows.sort((a, b) => this.timeValue(b) - this.timeValue(a) || a.title.localeCompare(b.title));
      } else if (this.sortKey === "title") {
        rows.sort((a, b) => a.title.localeCompare(b.title));
      } else {
        rows.sort((a, b) => (a.category || "").localeCompare(b.category || "") || a.title.localeCompare(b.title));
      }
      return rows;
    },
    pageCount() {
      return Math.max(1, Math.ceil(this.filteredItems.length / this.pageSize));
    },
    pagedItems() {
      const start = (this.page - 1) * this.pageSize;
      return this.filteredItems.slice(start, start + this.pageSize);
    },
    categoryLeaders() {
      const rows = this.categories.map((category) => {
        const items = this.items.filter((item) => item.category === category);
        return {
          category,
          items: items.length,
          assets: items.reduce((sum, item) => sum + Number(item.asset_count || 0), 0),
        };
      });
      return rows.sort((a, b) => b.items - a.items || a.category.localeCompare(b.category)).slice(0, 8);
    },
    categoryCards() {
      const rows = this.categories.map((category) => {
        const items = this.items.filter((item) => item.category === category);
        return {
          category,
          items: items.length,
          assets: items.reduce((sum, item) => sum + Number(item.asset_count || 0), 0),
        };
      });
      rows.sort((a, b) => b.items - a.items || a.category.localeCompare(b.category));
      return this.showAllCategories ? rows : rows.slice(0, 8);
    },
    popularItems() {
      return [...this.items].sort((a, b) => Number(b.asset_count || 0) - Number(a.asset_count || 0)).slice(0, 6);
    },
    recentItems() {
      return [...this.items].sort((a, b) => this.timeValue(b) - this.timeValue(a)).slice(0, 6);
    },
    selectedAssets() {
      if (!this.selectedItem) return [];
      return Object.entries(this.selectedItem.assets || {}).map(([key, asset]) => ({ key, ...asset }));
    },
    selectedProps() {
      return this.selectedItem?.properties || {};
    },
    temporalRows() {
      const props = this.selectedProps;
      return [
        ["Publication date", this.formatDate(props.datetime)],
        ["Temporal start", this.formatDate(props.start_datetime)],
        ["Temporal end", this.formatDate(props.end_datetime)],
      ].filter((row) => row[1]);
    },
    publicationRows() {
      const props = this.selectedProps;
      return [
        ["SEDAC DOI", props["ciesin:sedac_doi"]],
        ["NASA DOI", props["ciesin:nasa_doi"]],
        ["Harvard Dataverse DOI", props["ciesin:harvard_doi"]],
        ["License", props.license],
        ["Item status", props["ciesin:item_status"]],
        ["Source match", props["ciesin:source_match_status"]],
      ].filter((row) => row[1]);
    },
    geospatialRows() {
      const props = this.selectedProps;
      const bbox = this.selectedItem?.bbox || this.selected?.bbox;
      return [
        ["Bounding box", bbox ? bbox.map((value) => Number(value).toFixed(5)).join(", ") : ""],
        ["Geometry", this.selectedItem?.geometry?.type],
        ["EPSG", props["proj:epsg"]],
        ["Spatial representation", props["ciesin:spatial_representation"]],
        ["Data format", props["ciesin:data_format"]],
        ["Feature count", props["ciesin:feature_count"]],
      ].filter((row) => row[1] !== undefined && row[1] !== null && row[1] !== "");
    },
    dataStructureRows() {
      const props = this.selectedProps;
      const raster = props["raster:bands"]?.[0] || {};
      return [
        ["Field count", props["ciesin:field_count"]],
        ["Raster data type", raster.data_type],
        ["Bits per sample", raster.bits_per_sample],
        ["No data value", raster.nodata],
        ["Raster minimum", raster.statistics?.minimum],
        ["Raster maximum", raster.statistics?.maximum],
      ].filter((row) => row[1] !== undefined && row[1] !== null && row[1] !== "");
    },
    contactRows() {
      return this.selectedProps["ciesin:contacts"] || [];
    },
    selectedAssetSources() {
      return [...new Set(this.selectedAssets.map((asset) => asset["ciesin:source"]).filter(Boolean))].sort();
    },
    selectedAssetRoles() {
      const roles = new Set();
      this.selectedAssets.forEach((asset) => (asset.roles || []).forEach((role) => roles.add(role)));
      return [...roles].sort();
    },
    visibleAssets() {
      const q = this.assetQuery.toLowerCase();
      return this.selectedAssets.filter((asset) => {
        const haystack = [
          asset.key,
          asset.href,
          asset.title,
          asset.description,
          asset.type,
          asset["ciesin:source"],
          asset["ciesin:ac_doi"],
          ...(asset.roles || []),
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        return (
          (!q || haystack.includes(q)) &&
          (!this.assetSourceFilter || asset["ciesin:source"] === this.assetSourceFilter) &&
          (!this.assetRoleFilter || (asset.roles || []).includes(this.assetRoleFilter))
        );
      });
    },
    bboxLabel() {
      if (!this.selected || !this.selected.bbox) return "No bounding box available";
      return this.selected.bbox.map((value) => Number(value).toFixed(2)).join(", ");
    },
  },
  watch: {
    query() {
      this.resetPaging();
    },
    categoryFilter() {
      this.resetPaging();
    },
    statusFilter() {
      this.resetPaging();
    },
    sourceFilter() {
      this.resetPaging();
    },
    roleFilter() {
      this.resetPaging();
    },
    sortKey() {
      this.resetPaging();
    },
    pageCount(count) {
      if (this.page > count) this.page = count;
    },
  },
    async mounted() {
    try {
      const [itemsResponse, summaryResponse] = await Promise.all([
        fetch("./stac/items-index.json"),
        fetch("./stac/summary.json"),
      ]);
      this.items = await itemsResponse.json();
      this.summary = summaryResponse.ok ? await summaryResponse.json() : null;
      this.loadingMessage = "Loaded catalog index";
      const hashId = decodeURIComponent(window.location.hash.replace(/^#dataset=/, ""));
      if (hashId) {
        const item = this.items.find((row) => row.id === hashId);
        if (item) await this.selectItem(item, false);
      }
      window.addEventListener("hashchange", this.handleHashChange);
    } catch (error) {
      this.loadingMessage = `Could not load STAC index: ${error.message}`;
    }
  },
  methods: {
    async handleHashChange() {
      const hashId = decodeURIComponent(window.location.hash.replace(/^#dataset=/, ""));
      if (!hashId) {
        this.closeDetail(false);
        return;
      }
      if (this.selected?.id === hashId) return;
      const item = this.items.find((row) => row.id === hashId);
      if (item) await this.selectItem(item, false);
    },
    async selectItem(item, updateHash = true) {
      this.selected = item;
      this.assetQuery = "";
      this.assetSourceFilter = "";
      this.assetRoleFilter = "";
      this.loadingMessage = `Loading ${item.title}`;
      if (updateHash) {
        window.history.pushState(null, "", `#dataset=${encodeURIComponent(item.id)}`);
      }
      try {
        const response = await fetch(item.item_href);
        this.selectedItem = await response.json();
        await nextTick();
        this.drawMap();
        window.scrollTo({ top: 0, behavior: "smooth" });
      } catch (error) {
        this.selectedItem = null;
        this.loadingMessage = `Could not load item JSON: ${error.message}`;
      }
    },
    closeDetail(updateHash = true) {
      this.selected = null;
      this.selectedItem = null;
      if (updateHash && window.location.hash) {
        window.history.pushState(null, "", window.location.pathname);
      }
      window.scrollTo({ top: 0, behavior: "smooth" });
    },
    filterCategory(category) {
      this.categoryFilter = category;
      this.selected = null;
      this.selectedItem = null;
      nextTick(() => document.querySelector(".results-panel")?.scrollIntoView({ behavior: "smooth" }));
    },
    toggleAllCategories() {
      this.showAllCategories = !this.showAllCategories;
    },
    goHome() {
      this.selected = null;
      this.selectedItem = null;
      this.query = "";
      this.categoryFilter = "";
      this.statusFilter = "";
      this.sourceFilter = "";
      this.roleFilter = "";
      this.sortKey = "category";
      this.page = 1;
      if (window.location.hash) {
        window.history.pushState(null, "", window.location.pathname);
      }
      window.scrollTo({ top: 0, behavior: "smooth" });
    },
    showBrowse(sortKey) {
      this.sortKey = sortKey;
      this.selected = null;
      this.selectedItem = null;
      nextTick(() => document.querySelector(".results-panel")?.scrollIntoView({ behavior: "smooth" }));
    },
    resetPaging() {
      this.page = 1;
    },
    sourceName(value) {
      if (!value) return "source";
      return value
        .replace("nasa_cmr_granule", "NASA CMR")
        .replace("harvard_dataverse_file", "Harvard Dataverse")
        .replace("grid3_academic_commons", "GRID3 Academic Commons");
    },
    itemSourceLabel(item) {
      const sources = Object.entries(item.source_counts || {})
        .map(([source, count]) => `${this.sourceName(source)}: ${Number(count).toLocaleString()}`)
        .join(" · ");
      return sources || "No attached assets";
    },
    statusLabel(item) {
      return item.item_status || item.match_status || "status unavailable";
    },
    yearLabel(item) {
      const value = item.datetime || item.start_datetime || item.end_datetime || "";
      const year = String(value).match(/\d{4}/);
      return year ? year[0] : "No date";
    },
    formatDate(value) {
      if (value === null) return "Unknown";
      if (!value) return "";
      const parsed = Date.parse(value);
      if (Number.isNaN(parsed)) return value;
      return new Intl.DateTimeFormat("en", { year: "numeric", month: "short", day: "numeric" }).format(parsed);
    },
    doiHref(doi) {
      return `https://doi.org/${doi}`;
    },
    collectionHref(item) {
      return `./stac/collections/${item.collection}/collection.json`;
    },
    formatBytes(value) {
      const size = Number(value);
      if (!size) return "";
      const units = ["B", "KB", "MB", "GB", "TB"];
      let amount = size;
      let unit = 0;
      while (amount >= 1024 && unit < units.length - 1) {
        amount /= 1024;
        unit += 1;
      }
      return `${amount.toFixed(amount >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`;
    },
    timeValue(item) {
      const value = item.datetime || item.start_datetime || item.end_datetime || "";
      const parsed = Date.parse(value);
      return Number.isNaN(parsed) ? 0 : parsed;
    },
    drawMap() {
      const canvas = this.$refs.mapCanvas;
      if (!canvas) return;
      const ctx = canvas.getContext("2d");
      const width = canvas.width;
      const height = canvas.height;
      ctx.clearRect(0, 0, width, height);
      ctx.fillStyle = "#eef2ec";
      ctx.fillRect(0, 0, width, height);

      ctx.strokeStyle = "#cbd4cc";
      ctx.lineWidth = 1;
      for (let lon = -180; lon <= 180; lon += 60) {
        const x = ((lon + 180) / 360) * width;
        ctx.beginPath();
        ctx.moveTo(x, 0);
        ctx.lineTo(x, height);
        ctx.stroke();
      }
      for (let lat = -60; lat <= 60; lat += 30) {
        const y = ((90 - lat) / 180) * height;
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(width, y);
        ctx.stroke();
      }

      ctx.fillStyle = "#dce6df";
      ctx.fillRect(0, height * 0.28, width, height * 0.43);

      const bbox = this.selected && this.selected.bbox;
      if (!bbox) {
        ctx.fillStyle = "#65716b";
        ctx.font = "18px system-ui";
        ctx.fillText("No bbox", 24, 42);
        return;
      }

      const [west, south, east, north] = bbox.map(Number);
      const x1 = ((west + 180) / 360) * width;
      const x2 = ((east + 180) / 360) * width;
      const y1 = ((90 - north) / 180) * height;
      const y2 = ((90 - south) / 180) * height;
      ctx.fillStyle = "rgba(15, 111, 114, 0.26)";
      ctx.strokeStyle = "#0f6f72";
      ctx.lineWidth = 3;
      ctx.fillRect(x1, y1, Math.max(2, x2 - x1), Math.max(2, y2 - y1));
      ctx.strokeRect(x1, y1, Math.max(2, x2 - x1), Math.max(2, y2 - y1));
    },
  },
}).mount("#app");
