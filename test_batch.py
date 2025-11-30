import os
import pandas as pd
from main import process_batch
import shutil

def create_sample_excel(filename):
    df = pd.DataFrame({'Col1': [1, 2], 'Col2': ['A', 'B']})
    df.to_excel(filename, index=False)

def test_batch_processing():
    # Setup
    test_dir = "test_batch_output"
    if os.path.exists(test_dir):
        shutil.rmtree(test_dir)
    os.makedirs(test_dir)
    
    files = ["test1.xlsx", "test2.xlsx"]
    for f in files:
        create_sample_excel(f)
        
    # Test 1: Batch processing in same directory
    print("Testing batch processing (same dir)...")
    results = process_batch(files)
    assert len(results["success"]) == 2
    assert os.path.exists("test1.xml")
    assert os.path.exists("test2.xml")
    print("PASS")
    
    # Cleanup XMLs
    os.remove("test1.xml")
    os.remove("test2.xml")
    
    # Test 2: Batch processing with output directory
    print("Testing batch processing (custom output dir)...")
    results = process_batch(files, output_dir=test_dir)
    assert len(results["success"]) == 2
    assert os.path.exists(os.path.join(test_dir, "test1.xml"))
    assert os.path.exists(os.path.join(test_dir, "test2.xml"))
    print("PASS")
    
    # Cleanup
    for f in files:
        os.remove(f)
    shutil.rmtree(test_dir)
    print("All tests passed!")

if __name__ == "__main__":
    test_batch_processing()
