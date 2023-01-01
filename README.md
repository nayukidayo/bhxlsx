## 滨海文化中心热量表数据导出

compose.yaml

```yaml
services:
  xlsx:
    restart: always
    image: nayukidayo/xlsx:1.0
    logging:
      driver: local
    ports:
      - '51080:51080'
```
