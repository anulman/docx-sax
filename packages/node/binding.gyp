{
  "targets": [
    {
      "target_name": "docx_sax_node",
      "sources": ["src/addon.cc"],
      "cflags_cc": ["-std=c++17", "-fexceptions"],
      "conditions": [
        ["OS=='linux'", {
          "libraries": ["-ldl"]
        }]
      ]
    }
  ]
}
