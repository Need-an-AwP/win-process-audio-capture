{
  "targets": [
    {
      "target_name": "test_addon",
      "sources": [ "src/micAudioCapture_napi.cc" ],
      "include_dirs": [
        "C:\\Users\\17904\\Desktop\\voiceChat\\test\\src\\wil\\include",
        "C:\\Program Files (x86)\\Windows Kits\\10\\Lib\\10.0.22621.0\\um\\x64",
        "C:\\Program Files (x86)\\Windows Kits\\10\\Include\\10.0.22621.0\\um",
        "C:\\Users\\17904\Desktop\\voiceChat\\test\\node_modules\\node-addon-api",
        "C:\\Program Files (x86)\\Windows Kits\\10\\Include\\10.0.18362.0\\um"
      ],
      "libraries": [
          "mmdevapi.lib",
          "mf.lib",
          "mfplat.lib",
          "mfuuid.lib",
          "mfreadwrite.lib"
      ]
    }
  ]
}
