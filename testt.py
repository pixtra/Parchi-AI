import torch
from PIL import Image
from transformers import AutoModelForImageTextToText, AutoProcessor
from chandra.model.hf import generate_hf
from chandra.model.schema import BatchInputItem
from chandra.output import parse_markdown

model_id = "datalab-to/chandra-ocr-2"

# Load model, processor, and run inference
model = AutoModelForImageTextToText.from_pretrained(model_id, dtype=torch.bfloat16, device_map="auto")
model.eval()
model.processor = AutoProcessor.from_pretrained(model_id)
model.processor.tokenizer.padding_side = "left"

batch = [BatchInputItem(image=Image.open("/home/shahipaneer/WorkingFolder/Programming/Python/parchi-ai/test_docs/ewaybill.jpg"), prompt_type="ocr_layout")]
result = generate_hf(batch, model)[0]
print(parse_markdown(result.raw))
