## 根据sql生成数据库Word文档表格

自己懒得写文档，公司要求必须要如下格式的表格，所以自己写了一个用来使用，可以直接复制。


简单介绍下效果，通过数据库结构，正则提取数据后，将对应代码（通常mybatis生成对象名称为驼峰样式），备注根据是否有主键等获取到。


实际效果如下：

https://www.mccat.top/upload/image-tcTr.png

示例sql如下，navicat生成出来的ddl就能直接使用。

```sql
CREATE TABLE "public"."dm_template_feedback" (
  "id" int8 NOT NULL DEFAULT nextval('dm_template_feedback_id_seq'::regclass),
  "project_id" int8,
  "template_id" int8,
  "feedback_user_id" int8,
  "feedback_depart_id" int8,
  "feedback_time" timestamp(6),
  "content" text COLLATE "pg_catalog"."default",
  "attachment" text COLLATE "pg_catalog"."default",
  "creator_id" int8,
  "creator_name" varchar(255) COLLATE "pg_catalog"."default",
  "create_time" timestamp(6),
  CONSTRAINT "dm_template_feedback_pkey" PRIMARY KEY ("id")
)
;

ALTER TABLE "public"."dm_template_feedback" 
  OWNER TO "postgres";

COMMENT ON COLUMN "public"."dm_template_feedback"."project_id" IS '关联模型id';

COMMENT ON COLUMN "public"."dm_template_feedback"."template_id" IS '关联模型模版id';

COMMENT ON COLUMN "public"."dm_template_feedback"."feedback_user_id" IS '反馈人';

COMMENT ON COLUMN "public"."dm_template_feedback"."feedback_depart_id" IS '反馈单位';

COMMENT ON COLUMN "public"."dm_template_feedback"."feedback_time" IS '反馈时间';

COMMENT ON COLUMN "public"."dm_template_feedback"."content" IS '反馈内容';

COMMENT ON COLUMN "public"."dm_template_feedback"."attachment" IS '反馈附件';

COMMENT ON COLUMN "public"."dm_template_feedback"."creator_id" IS '创建人id';

COMMENT ON COLUMN "public"."dm_template_feedback"."creator_name" IS '创建人账号';

COMMENT ON COLUMN "public"."dm_template_feedback"."create_time" IS '创建时间';

COMMENT ON TABLE "public"."dm_template_feedback" IS '模型反馈表';
```