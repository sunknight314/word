import React from 'react';
import { Card, Row, Col, Typography, Tag, Descriptions, Space, Divider } from 'antd';
import { FileTextOutlined, SettingOutlined, FontSizeOutlined } from '@ant-design/icons';

const { Title, Text } = Typography;

const FormatDisplay = ({ formatData }) => {
  if (!formatData || !formatData.success) {
    return (
      <Card>
        <Text type="danger">格式识别失败</Text>
      </Card>
    );
  }

  const { display_data } = formatData;

  // 渲染页面设置
  const renderPageSettings = () => {
    const settings = display_data.page_settings;
    if (!settings) return null;

    const basicItems = settings.items.filter(item => item.category === 'basic');
    const marginItems = settings.items.filter(item => item.category === 'margins');

    return (
      <Card 
        title={
          <Space>
            <SettingOutlined />
            <span>{settings.title}</span>
          </Space>
        }
        style={{ marginBottom: 16 }}
      >
        <Descriptions bordered size="small" column={2}>
          {basicItems.map((item, index) => (
            <Descriptions.Item key={index} label={item.label}>
              <Tag color="blue">{item.value}</Tag>
            </Descriptions.Item>
          ))}
        </Descriptions>
        
        <Divider orientation="left">页边距设置</Divider>
        
        <Row gutter={[16, 16]}>
          {marginItems.map((item, index) => (
            <Col key={index} xs={24} sm={12} md={6}>
              <Card size="small" style={{ textAlign: 'center' }}>
                <Text type="secondary">{item.label}</Text>
                <br />
                <Text strong style={{ fontSize: 16 }}>{item.value}</Text>
              </Card>
            </Col>
          ))}
        </Row>
      </Card>
    );
  };

  // 渲染样式设置
  const renderStyles = () => {
    const styles = display_data.styles;
    if (!styles || styles.length === 0) return null;

    return (
      <Card 
        title={
          <Space>
            <FontSizeOutlined />
            <span>样式设置</span>
          </Space>
        }
        style={{ marginBottom: 16 }}
      >
        <Row gutter={[16, 16]}>
          {styles.map((style, index) => (
            <Col key={index} xs={24} md={12} lg={8}>
              <Card 
                type="inner" 
                title={style.name}
                extra={<Tag color="green">{style.key}</Tag>}
              >
                <Space direction="vertical" style={{ width: '100%' }}>
                  {/* 字体设置 */}
                  <div>
                    <Text strong>字体设置</Text>
                    <Descriptions size="small" column={1} style={{ marginTop: 8 }}>
                      {style.font.map((item, idx) => (
                        <Descriptions.Item key={idx} label={item.label}>
                          {item.value}
                        </Descriptions.Item>
                      ))}
                    </Descriptions>
                  </div>

                  {/* 段落设置 */}
                  <div>
                    <Text strong>段落设置</Text>
                    <Descriptions size="small" column={1} style={{ marginTop: 8 }}>
                      {style.paragraph.map((item, idx) => (
                        <Descriptions.Item key={idx} label={item.label}>
                          {item.value}
                        </Descriptions.Item>
                      ))}
                    </Descriptions>
                  </div>
                </Space>
              </Card>
            </Col>
          ))}
        </Row>
      </Card>
    );
  };

  // 渲染文档结构
  const renderDocumentStructure = () => {
    const structure = display_data.document_structure;
    if (!structure || !structure.items || structure.items.length === 0) return null;

    return (
      <Card 
        title={
          <Space>
            <FileTextOutlined />
            <span>{structure.title}</span>
          </Space>
        }
      >
        <Descriptions bordered size="small">
          {structure.items.map((item, index) => (
            <Descriptions.Item key={index} label={item.label} span={3}>
              <Tag color={item.category === 'toc' ? 'purple' : 'orange'}>
                {item.value}
              </Tag>
            </Descriptions.Item>
          ))}
        </Descriptions>
      </Card>
    );
  };

  return (
    <div style={{ padding: 24 }}>
      <Title level={2}>文档格式要求</Title>
      <Text type="secondary" style={{ marginBottom: 24, display: 'block' }}>
        以下是AI智能识别的格式配置，请仔细检查是否符合您的要求
      </Text>
      
      {renderPageSettings()}
      {renderStyles()}
      {renderDocumentStructure()}
    </div>
  );
};

export default FormatDisplay;