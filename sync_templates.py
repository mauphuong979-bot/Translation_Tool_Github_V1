import translation_lib as tl

if __name__ == "__main__":
    results = tl.sync_all_templates()
    for name, (success, msg) in results.items():
        if success:
            print(f"Success: Synced {name}")
        else:
            print(f"Error syncing {name}: {msg}")
    print("\nAll templates sync process completed!")

