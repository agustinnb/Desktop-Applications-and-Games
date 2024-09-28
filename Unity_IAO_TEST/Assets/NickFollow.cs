using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
public class NickFollow : MonoBehaviour
{
    public Vector3 pos;

    public GameObject robot;
    public Camera camera;
    public Text text;
    private Vector3 roboPos;
    private RectTransform rt;
    private RectTransform canvasRT;
    private Vector3 roboScreenPos;
    private string nickname;
    
    // Use this for initialization
    void Start()
    {
        roboPos = robot.transform.position;

        rt = GetComponent<RectTransform>();
        canvasRT = GetComponentInParent<Canvas>().GetComponent<RectTransform>();
        roboScreenPos = camera.WorldToViewportPoint(robot.transform.TransformPoint(roboPos));
        rt.anchorMax = roboScreenPos;
        rt.anchorMin = roboScreenPos;
        nickname = "Frixion";
        text.text = nickname;
    }

    // Update is called once per frame
    void Update()
    {
        roboScreenPos = camera.WorldToViewportPoint(robot.transform.TransformPoint(roboPos));
        rt.anchorMax = roboScreenPos;
        rt.anchorMin = roboScreenPos;
    }
}
