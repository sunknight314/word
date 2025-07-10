<template>
  <div class="format-display">
    <el-card class="header-card">
      <h1>文档格式要求</h1>
      <p class="subtitle">AI智能识别的格式配置</p>
    </el-card>

    <div v-if="!formatData || !formatData.success" class="error-container">
      <el-alert title="格式识别失败" type="error" :closable="false" />
    </div>

    <template v-else>
      <!-- 页面设置 -->
      <el-card v-if="displayData.page_settings" class="section-card">
        <template #header>
          <div class="card-header">
            <i class="el-icon-setting"></i>
            <span>{{ displayData.page_settings.title }}</span>
          </div>
        </template>

        <!-- 基本设置 -->
        <el-row :gutter="20">
          <el-col 
            v-for="item in basicSettings" 
            :key="item.label"
            :xs="24" :sm="12" :md="8"
          >
            <div class="setting-item">
              <span class="label">{{ item.label }}:</span>
              <el-tag type="info">{{ item.value }}</el-tag>
            </div>
          </el-col>
        </el-row>

        <!-- 页边距设置 -->
        <el-divider content-position="left">页边距设置</el-divider>
        <el-row :gutter="20">
          <el-col 
            v-for="item in marginSettings" 
            :key="item.label"
            :xs="12" :sm="6"
          >
            <el-card shadow="hover" class="margin-card">
              <div class="margin-label">{{ item.label }}</div>
              <div class="margin-value">{{ item.value }}</div>
            </el-card>
          </el-col>
        </el-row>
      </el-card>

      <!-- 样式设置 -->
      <el-card v-if="displayData.styles" class="section-card">
        <template #header>
          <div class="card-header">
            <i class="el-icon-document"></i>
            <span>样式设置</span>
          </div>
        </template>

        <el-collapse v-model="activeStyles">
          <el-collapse-item 
            v-for="(style, index) in displayData.styles" 
            :key="style.key"
            :title="style.name"
            :name="index"
          >
            <el-row :gutter="20">
              <!-- 字体设置 -->
              <el-col :xs="24" :md="12">
                <div class="style-group">
                  <h4>字体设置</h4>
                  <el-descriptions :column="1" size="small" border>
                    <el-descriptions-item 
                      v-for="item in style.font" 
                      :key="item.label"
                      :label="item.label"
                    >
                      <span :class="{ 'font-bold': item.value === '是' }">
                        {{ item.value }}
                      </span>
                    </el-descriptions-item>
                  </el-descriptions>
                </div>
              </el-col>

              <!-- 段落设置 -->
              <el-col :xs="24" :md="12">
                <div class="style-group">
                  <h4>段落设置</h4>
                  <el-descriptions :column="1" size="small" border>
                    <el-descriptions-item 
                      v-for="item in style.paragraph" 
                      :key="item.label"
                      :label="item.label"
                    >
                      {{ item.value }}
                    </el-descriptions-item>
                  </el-descriptions>
                </div>
              </el-col>
            </el-row>
          </el-collapse-item>
        </el-collapse>
      </el-card>

      <!-- 文档结构 -->
      <el-card 
        v-if="displayData.document_structure && displayData.document_structure.items.length > 0" 
        class="section-card"
      >
        <template #header>
          <div class="card-header">
            <i class="el-icon-files"></i>
            <span>{{ displayData.document_structure.title }}</span>
          </div>
        </template>

        <el-table :data="displayData.document_structure.items" stripe>
          <el-table-column prop="label" label="设置项" width="200" />
          <el-table-column prop="value" label="值">
            <template #default="{ row }">
              <el-tag :type="getTagType(row.category)">
                {{ row.value }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="category" label="类别" width="150">
            <template #default="{ row }">
              <span class="category-label">{{ getCategoryLabel(row.category) }}</span>
            </template>
          </el-table-column>
        </el-table>
      </el-card>
    </template>
  </div>
</template>

<script>
export default {
  name: 'FormatDisplay',
  props: {
    formatData: {
      type: Object,
      required: true
    }
  },
  data() {
    return {
      activeStyles: [0] // 默认展开第一个样式
    };
  },
  computed: {
    displayData() {
      return this.formatData?.display_data || {};
    },
    basicSettings() {
      return this.displayData.page_settings?.items.filter(item => item.category === 'basic') || [];
    },
    marginSettings() {
      return this.displayData.page_settings?.items.filter(item => item.category === 'margins') || [];
    }
  },
  methods: {
    getTagType(category) {
      const typeMap = {
        'toc': 'primary',
        'numbering': 'warning',
        'default': 'info'
      };
      return typeMap[category] || typeMap.default;
    },
    getCategoryLabel(category) {
      const labelMap = {
        'toc': '目录设置',
        'numbering': '页码设置'
      };
      return labelMap[category] || category;
    }
  }
};
</script>

<style scoped>
.format-display {
  padding: 20px;
  background-color: #f5f7fa;
  min-height: 100vh;
}

.header-card {
  margin-bottom: 20px;
  text-align: center;
}

.header-card h1 {
  margin: 0 0 10px 0;
  color: #303133;
}

.subtitle {
  color: #909399;
  margin: 0;
}

.section-card {
  margin-bottom: 20px;
}

.card-header {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 18px;
  font-weight: 600;
}

.card-header i {
  font-size: 20px;
  color: #409eff;
}

.setting-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px;
  background: #f4f4f5;
  border-radius: 4px;
  margin-bottom: 10px;
}

.setting-item .label {
  color: #606266;
  font-weight: 500;
}

.margin-card {
  text-align: center;
  cursor: pointer;
  transition: transform 0.2s;
}

.margin-card:hover {
  transform: translateY(-2px);
}

.margin-label {
  color: #909399;
  font-size: 14px;
  margin-bottom: 8px;
}

.margin-value {
  color: #303133;
  font-size: 18px;
  font-weight: 600;
}

.style-group {
  margin-bottom: 20px;
}

.style-group h4 {
  margin: 0 0 10px 0;
  color: #606266;
}

.font-bold {
  font-weight: 600;
  color: #409eff;
}

.category-label {
  font-size: 12px;
  color: #909399;
  text-transform: uppercase;
}

.error-container {
  padding: 40px;
  text-align: center;
}

/deep/ .el-collapse-item__header {
  font-size: 16px;
  font-weight: 600;
}

/deep/ .el-descriptions__label {
  width: 120px;
}
</style>